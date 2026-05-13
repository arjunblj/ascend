import {
	type CellStyle,
	cloneCellStyle,
	DEFAULT_STYLE_ID,
	parseA1,
	parseRange,
	type Sheet,
	type SheetComment,
	type SheetHyperlink,
	toA1,
	type Workbook,
} from '@ascend/core'
import { applyOperation } from '@ascend/engine'
import type { CellValue, InputValue, Operation, ScalarCellValue } from '@ascend/schema'
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

export type MutationJournalPreimage =
	| { readonly kind: 'cells'; readonly cells: readonly MutationJournalCellPreimage[] }
	| { readonly kind: 'comment'; readonly comment: MutationJournalCommentPreimage }
	| { readonly kind: 'hyperlink'; readonly hyperlink: MutationJournalHyperlinkPreimage }
	| { readonly kind: 'pane'; readonly pane: MutationJournalPanePreimage }

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
	const journalWorkbook = workbook.clone()
	const entries: MutationJournalEntry[] = []
	for (let opIndex = 0; opIndex < ops.length; opIndex++) {
		const op = ops[opIndex]
		if (!op) continue
		entries.push(buildJournalEntry(journalWorkbook, op, opIndex))
		const result = applyOperation(journalWorkbook, op)
		if (!result.ok) break
	}
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
		case 'setNumberFormat':
		case 'setStyle':
			return journalStyleRange(workbook, op, opIndex)
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
