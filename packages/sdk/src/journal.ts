import {
	type AutoFilter,
	type CellStyle,
	cloneCellStyle,
	DEFAULT_STYLE_ID,
	type DefinedName,
	type DefinedNameScope,
	parseA1,
	parseRange,
	type RangeRef,
	type Sheet,
	type SheetComment,
	type SheetConditionalFormat,
	type SheetConditionalFormatRule,
	type SheetDataValidation,
	type SheetHyperlink,
	type Table,
	type TableColumn,
	type TableStyleInfo,
	toA1,
	toRangeString,
	type Workbook,
} from '@ascend/core'
import { applyOperation } from '@ascend/engine'
import type {
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

export interface MutationJournalStructuralPreimage {
	readonly sheet: string
	readonly axis: 'row' | 'col'
	readonly at: number
	readonly count: number
	readonly deletedCells: readonly MutationJournalCellPreimage[]
}

export type MutationJournalPreimage =
	| { readonly kind: 'cells'; readonly cells: readonly MutationJournalCellPreimage[] }
	| { readonly kind: 'comment'; readonly comment: MutationJournalCommentPreimage }
	| { readonly kind: 'hyperlink'; readonly hyperlink: MutationJournalHyperlinkPreimage }
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
	| { readonly kind: 'defined-name'; readonly definedName: MutationJournalDefinedNamePreimage }
	| { readonly kind: 'table-rename'; readonly tableRename: MutationJournalTableRenamePreimage }
	| { readonly kind: 'table-column'; readonly tableColumn: MutationJournalTableColumnPreimage }
	| { readonly kind: 'table'; readonly table: MutationJournalTablePreimage }
	| { readonly kind: 'table-style'; readonly tableStyle: MutationJournalTableStylePreimage }
	| { readonly kind: 'structural'; readonly structural: MutationJournalStructuralPreimage }

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
		case 'setDefinedName':
			return journalSetDefinedName(workbook, op, opIndex)
		case 'deleteDefinedName':
			return journalDeleteDefinedName(workbook, op, opIndex)
		case 'renameTable':
			return journalRenameTable(op, opIndex)
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
		case 'freezePane':
			return journalFreezePane(workbook, op, opIndex)
		case 'renameSheet':
			return {
				opIndex,
				op,
				inverseOps: [{ op: 'renameSheet', sheet: op.newName, newName: op.sheet }],
				preimages: [],
				issues: [],
			}
		case 'addSheet':
			return {
				opIndex,
				op,
				inverseOps: [{ op: 'deleteSheet', sheet: op.name }],
				preimages: [],
				issues: [],
			}
		default:
			return null
	}
}

function journalSetCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setCells' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellPreimages(
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
	const cells = cellPreimages(workbook, op.sheet, [op.ref])
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
	const cells = cellPreimages(workbook, op.sheet, refs)
	if (op.what === 'styles') {
		return {
			opIndex,
			op,
			inverseOps: styleInverseOps(cells),
			preimages: [{ kind: 'cells', cells }],
			issues: [],
		}
	}
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

function journalRenameTable(
	op: Extract<Operation, { op: 'renameTable' }>,
	opIndex: number,
): DraftJournalEntry {
	const tableRename = { table: op.table, existed: true }
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'renameTable', table: op.newName, newName: op.table }],
		preimages: [{ kind: 'table-rename', tableRename }],
		issues: [],
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
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'table', table: preimage }],
		issues,
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
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'table', table: preimage }],
		issues,
	}
}

function journalSetTableColumn(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTableColumn' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = tableColumnPreimage(workbook, op.table, op.column)
	const { inverseOps, issues } = restoreTableColumnOps(op, preimage)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'table-column', tableColumn: preimage }],
		issues,
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
	return issues
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

function styleInverseOps(cells: readonly MutationJournalCellPreimage[]): Operation[] {
	return cells
		.filter((cell) => cell.existed)
		.map((cell) => ({
			op: 'setStyle',
			sheet: cell.sheet,
			range: cell.ref,
			style: cell.style,
		}))
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
