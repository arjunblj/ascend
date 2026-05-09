import { parseA1, parseRange } from '@ascend/core'
import type { Operation, StyleInput } from '@ascend/schema'
import { indexToColumn } from '@ascend/sdk'
import { selectionRef } from '../model/selection.ts'
import type { CommandContext } from '../runtime/types.ts'

type PageSetupInput = Extract<Operation, { op: 'setPageSetup' }>['setup']
type DataValidationRuleInput = Extract<Operation, { op: 'setDataValidation' }>['rule']
type ConditionalFormatRuleInput = Extract<Operation, { op: 'setConditionalFormat' }>['rule']

export type DialogId =
	| 'format-cells'
	| 'paste-special'
	| 'create-table'
	| 'sort'
	| 'filter'
	| 'find-replace'
	| 'data-validation'
	| 'conditional-formatting'
	| 'comment'
	| 'print-preview'
	| 'chart-wizard'
	| 'pivot-fields'

export interface DialogFieldDescriptor {
	readonly name: string
	readonly label: string
	readonly kind: 'text' | 'number' | 'boolean' | 'select' | 'range' | 'formula'
	readonly required?: boolean
	readonly options?: readonly string[]
}

export interface DialogDescriptor<TInput = unknown> {
	readonly id: DialogId
	readonly title: string
	readonly phase: 'foundation' | 'planned'
	readonly fields: readonly DialogFieldDescriptor[]
	defaultInput(ctx: CommandContext): TInput
	validate(input: TInput): readonly string[]
	toOperations(ctx: CommandContext, input: TInput): readonly Operation[]
}

export interface FormatCellsInput {
	readonly numberFormat?: string
	readonly bold?: boolean
	readonly italic?: boolean
	readonly horizontal?: 'general' | 'left' | 'center' | 'right'
}

export interface PasteSpecialInput {
	readonly source: string
	readonly target?: string
	readonly mode: 'all' | 'values' | 'formulas' | 'styles'
}

export interface SortInput {
	readonly range?: string
	readonly column: string | number
	readonly descending?: boolean
}

export interface CreateTableInput {
	readonly ref?: string
	readonly name: string
	readonly hasHeaders?: boolean
}

export interface FilterInput {
	readonly range?: string
}

export interface DataValidationInput {
	readonly range?: string
	readonly rule: DataValidationRuleInput
}

export interface ConditionalFormattingInput {
	readonly range?: string
	readonly rule: ConditionalFormatRuleInput
}

export interface CommentInput {
	readonly ref?: string
	readonly text: string
	readonly author?: string
}

export interface ChartWizardInput {
	readonly seriesIndex: number
	readonly sheet?: string
	readonly partPath?: string
	readonly chartIndex?: number
	readonly nameRef?: string
	readonly categoryRef?: string
	readonly valueRef?: string
}

export interface PivotFieldsInput {
	readonly cacheId?: number
	readonly partPath?: string
	readonly pivotTable?: string
	readonly sourceSheet?: string
	readonly sourceRef?: string
	readonly refreshOnLoad?: boolean
	readonly enableRefresh?: boolean
	readonly invalid?: boolean
	readonly saveData?: boolean
}

export interface PrintPreviewInput {
	readonly range?: string
	readonly orientation?: 'portrait' | 'landscape'
	readonly scale?: number
	readonly fitToWidth?: number
	readonly fitToHeight?: number
}

export interface FindReplaceInput {
	readonly range?: string
	readonly findText: string
	readonly replaceText?: string
	readonly action: 'find' | 'replace' | 'replaceAll'
	readonly lookIn: 'values' | 'formulas' | 'both'
	readonly matchCase?: boolean
	readonly matchEntireCell?: boolean
}

export const DIALOGS: readonly DialogDescriptor[] = [
	{
		id: 'format-cells',
		title: 'Format Cells',
		phase: 'foundation',
		fields: [
			{ name: 'numberFormat', label: 'Number format', kind: 'text' },
			{ name: 'bold', label: 'Bold', kind: 'boolean' },
			{ name: 'italic', label: 'Italic', kind: 'boolean' },
			{
				name: 'horizontal',
				label: 'Horizontal alignment',
				kind: 'select',
				options: ['general', 'left', 'center', 'right'],
			},
		],
		defaultInput: () => ({}) satisfies FormatCellsInput,
		validate: (input) => {
			const typed = input as FormatCellsInput
			if (
				typed.horizontal !== undefined &&
				!['general', 'left', 'center', 'right'].includes(typed.horizontal)
			) {
				return ['Horizontal alignment must be general, left, center, or right.']
			}
			return []
		},
		toOperations(ctx, input) {
			const typed = input as FormatCellsInput
			const style: StyleInput = {
				...(typed.numberFormat ? { numberFormat: typed.numberFormat } : {}),
				...(typed.bold !== undefined || typed.italic !== undefined
					? {
							font: {
								...(typed.bold !== undefined ? { bold: typed.bold } : {}),
								...(typed.italic !== undefined ? { italic: typed.italic } : {}),
							},
						}
					: {}),
				...(typed.horizontal ? { alignment: { horizontal: typed.horizontal } } : {}),
			}
			if (isEmptyStyle(style)) return []
			return [
				{
					op: 'setStyle',
					sheet: ctx.sheet,
					range: selectionRef(ctx.selection, indexToColumn),
					style,
				},
			]
		},
	},
	{
		id: 'paste-special',
		title: 'Paste Special',
		phase: 'foundation',
		fields: [
			{ name: 'source', label: 'Source range', kind: 'range', required: true },
			{ name: 'target', label: 'Target cell', kind: 'range' },
			{
				name: 'mode',
				label: 'Paste mode',
				kind: 'select',
				options: ['all', 'values', 'formulas', 'styles'],
			},
		],
		defaultInput: (ctx) =>
			({
				source: selectionRef(ctx.selection, indexToColumn),
				mode: 'all',
			}) satisfies PasteSpecialInput,
		validate: (input) => {
			const typed = input as PasteSpecialInput
			return [
				...validateRangeText(typed.source, 'Source range', true),
				...validateRangeText(typed.target, 'Target cell'),
				...validateEnum(typed.mode, 'Paste mode', ['all', 'values', 'formulas', 'styles']),
			]
		},
		toOperations(ctx, input) {
			const typed = input as PasteSpecialInput
			return [
				{
					op: 'copyRange',
					sheet: ctx.sheet,
					source: typed.source,
					target: typed.target ?? selectionRef(ctx.selection, indexToColumn),
					mode: typed.mode,
				},
			]
		},
	},
	{
		id: 'sort',
		title: 'Sort',
		phase: 'foundation',
		fields: [
			{ name: 'range', label: 'Range', kind: 'range' },
			{ name: 'column', label: 'Sort by', kind: 'text', required: true },
			{ name: 'descending', label: 'Descending', kind: 'boolean' },
		],
		defaultInput: (ctx) =>
			({
				range: selectionRef(ctx.selection, indexToColumn),
				column: indexToColumn(ctx.selection.active.col),
			}) satisfies SortInput,
		validate: (input) => {
			const typed = input as SortInput
			return [...validateRangeText(typed.range, 'Sort range'), ...validateSortColumn(typed.column)]
		},
		toOperations(ctx, input) {
			const typed = input as SortInput
			const sortBy = {
				column: typed.column,
				...(typed.descending !== undefined ? { descending: typed.descending } : {}),
			}
			return [
				{
					op: 'sortRange',
					sheet: ctx.sheet,
					range: typed.range ?? selectionRef(ctx.selection, indexToColumn),
					by: [sortBy],
				},
			]
		},
	},
	{
		id: 'create-table',
		title: 'Create Table',
		phase: 'foundation',
		fields: [
			{ name: 'ref', label: 'Table range', kind: 'range' },
			{ name: 'name', label: 'Table name', kind: 'text', required: true },
			{ name: 'hasHeaders', label: 'My table has headers', kind: 'boolean' },
		],
		defaultInput: (ctx) =>
			({
				ref: selectionRef(ctx.selection, indexToColumn),
				name: `Table${ctx.selection.active.row + 1}_${ctx.selection.active.col + 1}`,
				hasHeaders: true,
			}) satisfies CreateTableInput,
		validate: (input) => {
			const typed = input as CreateTableInput
			const errors: string[] = []
			errors.push(...validateRangeText(typed.ref, 'Table range'))
			if (!typed.name?.trim()) errors.push('Table name is required.')
			if (/\s/.test(typed.name ?? '')) errors.push('Table name cannot contain spaces.')
			return errors
		},
		toOperations(ctx, input) {
			const typed = input as CreateTableInput
			return [
				{
					op: 'createTable',
					sheet: ctx.sheet,
					ref: typed.ref ?? selectionRef(ctx.selection, indexToColumn),
					name: typed.name,
					hasHeaders: typed.hasHeaders ?? true,
				},
			]
		},
	},
	{
		id: 'filter',
		title: 'Filter',
		phase: 'foundation',
		fields: [{ name: 'range', label: 'Range', kind: 'range' }],
		defaultInput: (ctx) =>
			({ range: selectionRef(ctx.selection, indexToColumn) }) satisfies FilterInput,
		validate: (input) => validateRangeText((input as FilterInput).range, 'Filter range'),
		toOperations(ctx, input) {
			const typed = input as FilterInput
			return [
				{
					op: 'setAutoFilter',
					sheet: ctx.sheet,
					range: typed.range ?? selectionRef(ctx.selection, indexToColumn),
				},
			]
		},
	},
	{
		id: 'data-validation',
		title: 'Data Validation',
		phase: 'foundation',
		fields: [
			{ name: 'range', label: 'Range', kind: 'range' },
			{ name: 'formula1', label: 'Formula 1', kind: 'formula', required: true },
		],
		defaultInput: (ctx) =>
			({
				range: selectionRef(ctx.selection, indexToColumn),
				rule: { type: 'list', formula1: '""' },
			}) satisfies DataValidationInput,
		validate: (input) => {
			const typed = input as DataValidationInput
			return [
				...validateRangeText(typed.range, 'Validation range'),
				...validateDataValidationRule(typed.rule),
			]
		},
		toOperations(ctx, input) {
			const typed = input as DataValidationInput
			return [
				{
					op: 'setDataValidation',
					sheet: ctx.sheet,
					range: typed.range ?? selectionRef(ctx.selection, indexToColumn),
					rule: typed.rule,
				},
			]
		},
	},
	{
		id: 'conditional-formatting',
		title: 'Conditional Formatting',
		phase: 'foundation',
		fields: [
			{ name: 'range', label: 'Range', kind: 'range' },
			{ name: 'formula', label: 'Formula', kind: 'formula', required: true },
		],
		defaultInput: (ctx) =>
			({
				range: selectionRef(ctx.selection, indexToColumn),
				rule: { type: 'expression', formula: 'TRUE' },
			}) satisfies ConditionalFormattingInput,
		validate: (input) => {
			const typed = input as ConditionalFormattingInput
			return [
				...validateRangeText(typed.range, 'Conditional format range'),
				...validateConditionalFormatRule(typed.rule),
			]
		},
		toOperations(ctx, input) {
			const typed = input as ConditionalFormattingInput
			return [
				{
					op: 'setConditionalFormat',
					sheet: ctx.sheet,
					range: typed.range ?? selectionRef(ctx.selection, indexToColumn),
					rule: typed.rule,
				},
			]
		},
	},
	{
		id: 'comment',
		title: 'New Comment',
		phase: 'foundation',
		fields: [
			{ name: 'ref', label: 'Cell', kind: 'range' },
			{ name: 'text', label: 'Comment', kind: 'text', required: true },
			{ name: 'author', label: 'Author', kind: 'text' },
		],
		defaultInput: (ctx) =>
			({
				ref: activeCellRef(ctx),
				text: '',
				author: 'Ascend',
			}) satisfies CommentInput,
		validate: (input) => {
			const typed = input as CommentInput
			const errors = validateCellText(typed.ref, 'Comment cell')
			if (!typed.text?.trim()) errors.push('Comment text is required.')
			return errors
		},
		toOperations(ctx, input) {
			const typed = input as CommentInput
			return [
				{
					op: 'setComment',
					sheet: ctx.sheet,
					ref: typed.ref ?? activeCellRef(ctx),
					text: typed.text,
					...(typed.author ? { author: typed.author } : {}),
				},
			]
		},
	},
	{
		id: 'find-replace',
		title: 'Find and Replace',
		phase: 'foundation',
		fields: [
			{ name: 'range', label: 'Range', kind: 'range' },
			{ name: 'findText', label: 'Find what', kind: 'text', required: true },
			{ name: 'replaceText', label: 'Replace with', kind: 'text' },
			{
				name: 'action',
				label: 'Action',
				kind: 'select',
				options: ['find', 'replace', 'replaceAll'],
			},
			{
				name: 'lookIn',
				label: 'Look in',
				kind: 'select',
				options: ['values', 'formulas', 'both'],
			},
			{ name: 'matchCase', label: 'Match case', kind: 'boolean' },
			{ name: 'matchEntireCell', label: 'Match entire cell', kind: 'boolean' },
		],
		defaultInput: (ctx) =>
			({
				range: selectionRef(ctx.selection, indexToColumn),
				findText: '',
				replaceText: '',
				action: 'find',
				lookIn: 'values',
				matchCase: false,
				matchEntireCell: false,
			}) satisfies FindReplaceInput,
		validate: (input) => {
			const typed = input as FindReplaceInput
			const errors: string[] = []
			if (!typed.findText) errors.push('Find text is required.')
			if (!['find', 'replace', 'replaceAll'].includes(typed.action)) {
				errors.push('Action must be find, replace, or replaceAll.')
			}
			if (!['values', 'formulas', 'both'].includes(typed.lookIn)) {
				errors.push('Look in must be values, formulas, or both.')
			}
			return errors
		},
		toOperations: () => [],
	},
	{
		id: 'chart-wizard',
		title: 'Chart Wizard',
		phase: 'foundation',
		fields: [
			{ name: 'seriesIndex', label: 'Series index', kind: 'number', required: true },
			{ name: 'sheet', label: 'Chart sheet', kind: 'text' },
			{ name: 'partPath', label: 'Chart part path', kind: 'text' },
			{ name: 'chartIndex', label: 'Chart index', kind: 'number' },
			{ name: 'nameRef', label: 'Series name ref', kind: 'formula' },
			{ name: 'categoryRef', label: 'Category ref', kind: 'formula' },
			{ name: 'valueRef', label: 'Values ref', kind: 'formula' },
		],
		defaultInput: (ctx) =>
			({
				seriesIndex: 0,
				sheet: ctx.sheet,
				valueRef: `${ctx.sheet}!${selectionRef(ctx.selection, indexToColumn)}`,
			}) satisfies ChartWizardInput,
		validate: (input) => {
			const typed = input as ChartWizardInput
			const errors: string[] = []
			if (!Number.isInteger(typed.seriesIndex) || typed.seriesIndex < 0) {
				errors.push('Series index must be a non-negative integer.')
			}
			if (!typed.nameRef && !typed.categoryRef && !typed.valueRef) {
				errors.push('At least one chart source reference is required.')
			}
			return errors
		},
		toOperations(_ctx, input) {
			const typed = input as ChartWizardInput
			return [
				{
					op: 'setChartSeriesSource',
					seriesIndex: typed.seriesIndex,
					...(typed.sheet ? { sheet: typed.sheet } : {}),
					...(typed.partPath ? { partPath: typed.partPath } : {}),
					...(typed.chartIndex !== undefined ? { chartIndex: typed.chartIndex } : {}),
					...(typed.nameRef ? { nameRef: typed.nameRef } : {}),
					...(typed.categoryRef ? { categoryRef: typed.categoryRef } : {}),
					...(typed.valueRef ? { valueRef: typed.valueRef } : {}),
				},
			]
		},
	},
	{
		id: 'pivot-fields',
		title: 'Pivot Fields',
		phase: 'foundation',
		fields: [
			{ name: 'cacheId', label: 'Cache id', kind: 'number' },
			{ name: 'partPath', label: 'Cache part path', kind: 'text' },
			{ name: 'pivotTable', label: 'Pivot table', kind: 'text' },
			{ name: 'sourceSheet', label: 'Source sheet', kind: 'text' },
			{ name: 'sourceRef', label: 'Source range', kind: 'range' },
			{ name: 'refreshOnLoad', label: 'Refresh on open', kind: 'boolean' },
			{ name: 'enableRefresh', label: 'Enable refresh', kind: 'boolean' },
			{ name: 'invalid', label: 'Mark cache stale', kind: 'boolean' },
			{ name: 'saveData', label: 'Save cached data', kind: 'boolean' },
		],
		defaultInput: (ctx) =>
			({
				sourceSheet: ctx.sheet,
				sourceRef: selectionRef(ctx.selection, indexToColumn),
				refreshOnLoad: true,
				invalid: true,
			}) satisfies PivotFieldsInput,
		validate: (input) => {
			const typed = input as PivotFieldsInput
			const errors: string[] = []
			if (typed.cacheId === undefined && !typed.partPath && !typed.pivotTable) {
				errors.push('Pivot cache id, part path, or pivot table is required.')
			}
			if (
				typed.sourceSheet === undefined &&
				typed.sourceRef === undefined &&
				typed.refreshOnLoad === undefined &&
				typed.enableRefresh === undefined &&
				typed.invalid === undefined &&
				typed.saveData === undefined
			) {
				errors.push('At least one pivot cache update is required.')
			}
			return errors
		},
		toOperations(_ctx, input) {
			const typed = input as PivotFieldsInput
			return [
				{
					op: 'setPivotCache',
					...(typed.cacheId !== undefined ? { cacheId: typed.cacheId } : {}),
					...(typed.partPath ? { partPath: typed.partPath } : {}),
					...(typed.pivotTable ? { pivotTable: typed.pivotTable } : {}),
					...(typed.sourceSheet ? { sourceSheet: typed.sourceSheet } : {}),
					...(typed.sourceRef ? { sourceRef: typed.sourceRef } : {}),
					...(typed.refreshOnLoad !== undefined ? { refreshOnLoad: typed.refreshOnLoad } : {}),
					...(typed.enableRefresh !== undefined ? { enableRefresh: typed.enableRefresh } : {}),
					...(typed.invalid !== undefined ? { invalid: typed.invalid } : {}),
					...(typed.saveData !== undefined ? { saveData: typed.saveData } : {}),
				},
			]
		},
	},
	{
		id: 'print-preview',
		title: 'Print Preview',
		phase: 'foundation',
		fields: [
			{ name: 'range', label: 'Print area', kind: 'range' },
			{
				name: 'orientation',
				label: 'Orientation',
				kind: 'select',
				options: ['portrait', 'landscape'],
			},
			{ name: 'scale', label: 'Scale percent', kind: 'number' },
			{ name: 'fitToWidth', label: 'Fit to pages wide', kind: 'number' },
			{ name: 'fitToHeight', label: 'Fit to pages tall', kind: 'number' },
		],
		defaultInput: (ctx) =>
			({
				range: selectionRef(ctx.selection, indexToColumn),
				orientation: 'landscape',
				fitToWidth: 1,
				fitToHeight: 0,
			}) satisfies PrintPreviewInput,
		validate: (input) => {
			const typed = input as PrintPreviewInput
			const errors: string[] = [...validateRangeText(typed.range, 'Print area')]
			for (const [label, value] of [
				['Scale percent', typed.scale],
				['Fit to pages wide', typed.fitToWidth],
				['Fit to pages tall', typed.fitToHeight],
			] as const) {
				if (value !== undefined && (!Number.isInteger(value) || value < 0)) {
					errors.push(`${label} must be a non-negative integer.`)
				}
			}
			return errors
		},
		toOperations(ctx, input) {
			const typed = input as PrintPreviewInput
			const setup: PageSetupInput = {
				...(typed.orientation ? { orientation: typed.orientation } : {}),
				...(typed.scale !== undefined ? { scale: typed.scale } : {}),
				...(typed.fitToWidth !== undefined ? { fitToWidth: typed.fitToWidth } : {}),
				...(typed.fitToHeight !== undefined ? { fitToHeight: typed.fitToHeight } : {}),
			}
			return [
				{
					op: 'setPrintArea',
					sheet: ctx.sheet,
					range: typed.range ?? selectionRef(ctx.selection, indexToColumn),
				},
				{ op: 'setPageSetup', sheet: ctx.sheet, setup },
			]
		},
	},
]

export function findDialog(id: DialogId): DialogDescriptor | undefined {
	return DIALOGS.find((dialog) => dialog.id === id)
}

export function buildDialogOperations(
	id: DialogId,
	ctx: CommandContext,
	input?: unknown,
): readonly Operation[] {
	const dialog = findDialog(id)
	if (!dialog) throw new Error(`Unknown dialog: ${id}`)
	const resolvedInput = input ?? dialog.defaultInput(ctx)
	const errors = dialog.validate(resolvedInput)
	if (errors.length > 0) throw new Error(errors.join('; '))
	return dialog.toOperations(ctx, resolvedInput)
}

function activeCellRef(ctx: CommandContext): string {
	return `${indexToColumn(ctx.selection.active.col)}${ctx.selection.active.row + 1}`
}

function isEmptyStyle(style: StyleInput): boolean {
	return (
		style.numberFormat === undefined &&
		style.font === undefined &&
		style.fill === undefined &&
		style.border === undefined &&
		style.alignment === undefined &&
		style.protection === undefined
	)
}

function validateRangeText(value: string | undefined, label: string, required = false): string[] {
	if (value === undefined || value === '') return required ? [`${label} is required.`] : []
	if (typeof value !== 'string') return [`${label} must be a range.`]
	try {
		parseRange(value)
		return []
	} catch {
		return [`${label} must be a valid A1 range.`]
	}
}

function validateCellText(value: string | undefined, label: string): string[] {
	if (value === undefined || value === '') return []
	if (typeof value !== 'string') return [`${label} must be a cell reference.`]
	try {
		parseA1(value)
		return []
	} catch {
		return [`${label} must be a single A1 cell reference.`]
	}
}

function validateSortColumn(value: string | number | undefined): string[] {
	if (typeof value === 'number') {
		return Number.isInteger(value) && value >= 0
			? []
			: ['Sort column must be a non-negative integer.']
	}
	if (typeof value === 'string' && value.trim() !== '') return []
	return ['Sort column is required.']
}

function validateEnum<T extends string>(
	value: T | undefined,
	label: string,
	allowed: readonly T[],
): string[] {
	if (value === undefined || allowed.includes(value)) return []
	return [`${label} must be one of: ${allowed.join(', ')}.`]
}

function validateDataValidationRule(rule: DataValidationRuleInput | undefined): string[] {
	if (!isRecord(rule)) return ['Validation rule is required.']
	const type = typeof rule.type === 'string' ? rule.type : undefined
	const errors = type
		? validateEnum(type, 'Validation type', [
				'list',
				'whole',
				'decimal',
				'date',
				'time',
				'textLength',
				'custom',
			] as const)
		: ['Validation type must be one of: list, whole, decimal, date, time, textLength, custom.']
	if (!rule.formula1) errors.push('Validation formula is required.')
	return errors
}

function validateConditionalFormatRule(rule: ConditionalFormatRuleInput | undefined): string[] {
	if (!isRecord(rule)) return ['Conditional format rule is required.']
	const type = typeof rule.type === 'string' ? rule.type : undefined
	const errors = type
		? validateEnum(type, 'Conditional format type', [
				'cellIs',
				'expression',
				'colorScale',
				'dataBar',
				'iconSet',
				'top10',
				'aboveAverage',
				'duplicateValues',
				'containsText',
			] as const)
		: [
				'Conditional format type must be one of: cellIs, expression, colorScale, dataBar, iconSet, top10, aboveAverage, duplicateValues, containsText.',
			]
	if (
		(rule.type === 'cellIs' || rule.type === 'expression' || rule.type === 'containsText') &&
		!rule.formula
	) {
		errors.push('Conditional format formula is required.')
	}
	return errors
}

function isRecord(value: unknown): value is Record<string, unknown> {
	return typeof value === 'object' && value !== null
}
