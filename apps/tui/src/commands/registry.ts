import type { Operation } from '@ascend/schema'
import { indexToColumn } from '@ascend/sdk'
import type { DialogId } from '../dialogs/index.ts'
import { buildDialogOperations } from '../dialogs/index.ts'
import { selectionRef } from '../model/selection.ts'
import type { CommandContext, CommandDescriptor } from '../runtime/types.ts'

const noOps = (): readonly Operation[] => []

export const COMMAND_REGISTRY: readonly CommandDescriptor[] = [
	command('file.open', 'Open workbook', 'file', ['Ctrl+O'], [':open'], ['fileHub'], noOps),
	command('file.save', 'Save workbook', 'file', ['Ctrl+S'], [':save'], ['grid'], noOps),
	command('file.saveAs', 'Save As', 'file', ['F12'], [':save-as'], ['grid'], noOps),
	command('file.export', 'Export', 'file', ['Ctrl+Shift+S'], [':export'], ['grid'], noOps),
	command(
		'file.printPreview',
		'Print Preview',
		'file',
		[],
		[':print'],
		['grid'],
		(ctx, input) => buildDialogOperations('print-preview', ctx, input),
		'print-preview',
	),
	command('home.copy', 'Copy', 'home', ['Ctrl+C'], [':copy'], ['grid'], noOps),
	command('home.cut', 'Cut', 'home', ['Ctrl+X'], [':cut'], ['grid'], noOps),
	command('home.paste', 'Paste', 'home', ['Ctrl+V'], [':paste'], ['grid'], noOps),
	command(
		'home.pasteSpecial',
		'Paste Special',
		'home',
		['Ctrl+Alt+V'],
		[':paste values'],
		['grid'],
		(ctx, input) => buildDialogOperations('paste-special', ctx, input),
		'paste-special',
	),
	command('home.clear', 'Clear contents', 'home', ['Delete'], [':clear'], ['grid'], (ctx) => [
		{
			op: 'clearRange',
			sheet: ctx.sheet,
			range: selectionRef(ctx.selection, indexToColumn),
			what: 'all',
		},
	]),
	command(
		'home.formatCells',
		'Format Cells',
		'home',
		['Ctrl+1'],
		[':format'],
		['grid'],
		(ctx, input) => buildDialogOperations('format-cells', ctx, input),
		'format-cells',
	),
	command(
		'home.findReplace',
		'Find and Replace',
		'home',
		['Ctrl+F', 'Ctrl+H'],
		[':find', ':replace'],
		['grid'],
		(ctx, input) => buildDialogOperations('find-replace', ctx, input),
		'find-replace',
	),
	command(
		'insert.table',
		'Create table',
		'insert',
		['Ctrl+T'],
		[':table create'],
		['grid'],
		(ctx, input) => buildDialogOperations('create-table', ctx, input),
		'create-table',
	),
	command(
		'insert.chart',
		'Create chart',
		'insert',
		['Alt+F1'],
		[':chart'],
		['grid'],
		(ctx, input) => buildDialogOperations('chart-wizard', ctx, input),
		'chart-wizard',
	),
	command('formulas.autosum', 'AutoSum', 'formulas', ['Alt+='], [':autosum'], ['grid'], noOps),
	command('formulas.recalculate', 'Recalculate', 'formulas', ['F9'], [':recalc'], ['grid'], noOps),
	command(
		'formulas.tracePrecedents',
		'Trace precedents',
		'formulas',
		[],
		[':trace precedents'],
		['grid'],
		noOps,
	),
	command(
		'data.filter',
		'Toggle filter',
		'data',
		['Ctrl+Shift+L'],
		[':filter'],
		['grid'],
		(ctx, input) => buildDialogOperations('filter', ctx, input),
	),
	command(
		'data.sort',
		'Sort',
		'data',
		[],
		[':sort'],
		['grid'],
		(ctx, input) => buildDialogOperations('sort', ctx, input),
		'sort',
	),
	command(
		'data.validation',
		'Data Validation',
		'data',
		[],
		[':validate'],
		['grid'],
		(ctx, input) => buildDialogOperations('data-validation', ctx, input),
		'data-validation',
	),
	command(
		'data.conditionalFormatting',
		'Conditional Formatting',
		'data',
		[],
		[':conditional-format'],
		['grid'],
		(ctx, input) => buildDialogOperations('conditional-formatting', ctx, input),
		'conditional-formatting',
	),
	command(
		'data.pivotFields',
		'Pivot Fields',
		'data',
		[],
		[':pivot'],
		['grid'],
		(ctx, input) => buildDialogOperations('pivot-fields', ctx, input),
		'pivot-fields',
	),
	command(
		'review.comment',
		'New comment',
		'review',
		['Shift+F2'],
		[':comment'],
		['grid'],
		(ctx, input) => buildDialogOperations('comment', ctx, input),
		'comment',
	),
	command('view.freeze', 'Freeze panes', 'view', [], [':freeze'], ['grid'], noOps),
	command(
		'view.showFormulas',
		'Show formulas',
		'view',
		['Ctrl+`'],
		[':show formulas'],
		['grid'],
		noOps,
	),
	command(
		'view.objects',
		'Object Inspector',
		'view',
		[],
		[':objects', ':charts', ':pivots'],
		['grid', 'inspector'],
		noOps,
	),
]

export function findCommand(query: string): CommandDescriptor | undefined {
	const normalized = query.trim().toLowerCase()
	return COMMAND_REGISTRY.find(
		(command) =>
			command.id.toLowerCase() === normalized ||
			command.title.toLowerCase() === normalized ||
			command.fallbackKeys.some((key) => key.slice(1).toLowerCase() === normalized),
	)
}

export function listCommands(): readonly CommandDescriptor[] {
	return COMMAND_REGISTRY
}

export function commandsForGroup(group: CommandDescriptor['group']): readonly CommandDescriptor[] {
	return COMMAND_REGISTRY.filter((command) => command.group === group)
}

function command(
	id: string,
	title: string,
	group: CommandDescriptor['group'],
	excelKeys: readonly string[],
	fallbackKeys: readonly string[],
	contexts: CommandDescriptor['contexts'],
	toOperations: (ctx: CommandContext, input: unknown) => readonly Operation[],
	dialogId?: DialogId,
): CommandDescriptor {
	return {
		id,
		title,
		group,
		excelKeys,
		fallbackKeys,
		contexts,
		toOperations,
		...(dialogId ? { dialogId } : {}),
	}
}
