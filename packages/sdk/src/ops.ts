import type { InputValue, Operation } from '@ascend/schema'

export interface OperationSchema {
	readonly op: string
	readonly description: string
	readonly requiredFields: readonly string[]
	readonly optionalFields?: readonly string[]
}

export function listOperations(): readonly OperationSchema[] {
	return [
		{ op: 'setCells', description: 'Set cell values', requiredFields: ['sheet', 'updates'] },
		{
			op: 'setFormula',
			description: 'Set a formula in a cell',
			requiredFields: ['sheet', 'ref', 'formula'],
		},
		{
			op: 'fillFormula',
			description: 'Fill a formula across a range',
			requiredFields: ['sheet', 'range', 'formula'],
		},
		{
			op: 'clearRange',
			description: 'Clear values, formulas, styles, or all from a range',
			requiredFields: ['sheet', 'range', 'what'],
		},
		{ op: 'insertRows', description: 'Insert rows', requiredFields: ['sheet', 'at', 'count'] },
		{ op: 'deleteRows', description: 'Delete rows', requiredFields: ['sheet', 'at', 'count'] },
		{ op: 'insertCols', description: 'Insert columns', requiredFields: ['sheet', 'at', 'count'] },
		{ op: 'deleteCols', description: 'Delete columns', requiredFields: ['sheet', 'at', 'count'] },
		{
			op: 'addSheet',
			description: 'Add a new sheet',
			requiredFields: ['name'],
			optionalFields: ['position'],
		},
		{ op: 'deleteSheet', description: 'Delete a sheet', requiredFields: ['sheet'] },
		{ op: 'renameSheet', description: 'Rename a sheet', requiredFields: ['sheet', 'newName'] },
		{
			op: 'moveSheet',
			description: 'Move a sheet to a new position',
			requiredFields: ['sheet', 'position'],
		},
		{
			op: 'createTable',
			description: 'Create a table from a range',
			requiredFields: ['sheet', 'ref', 'name', 'hasHeaders'],
		},
		{ op: 'appendRows', description: 'Append rows to a table', requiredFields: ['table', 'rows'] },
		{
			op: 'sortRange',
			description: 'Sort a range by columns',
			requiredFields: ['sheet', 'range', 'by'],
		},
		{ op: 'mergeCells', description: 'Merge cells in a range', requiredFields: ['sheet', 'range'] },
		{
			op: 'unmergeCells',
			description: 'Unmerge cells in a range',
			requiredFields: ['sheet', 'range'],
		},
		{
			op: 'setColWidth',
			description: 'Set column width',
			requiredFields: ['sheet', 'col', 'width'],
		},
		{
			op: 'setRowHeight',
			description: 'Set row height',
			requiredFields: ['sheet', 'row', 'height'],
		},
		{
			op: 'setComment',
			description: 'Set a cell comment',
			requiredFields: ['sheet', 'ref', 'text'],
			optionalFields: ['author'],
		},
		{
			op: 'setHyperlink',
			description: 'Set a cell hyperlink',
			requiredFields: ['sheet', 'ref', 'url'],
			optionalFields: ['display'],
		},
		{
			op: 'setNumberFormat',
			description: 'Set number format for a range',
			requiredFields: ['sheet', 'range', 'format'],
		},
		{
			op: 'setDefinedName',
			description: 'Define a named range or formula',
			requiredFields: ['name', 'ref'],
			optionalFields: ['scope'],
		},
		{
			op: 'deleteDefinedName',
			description: 'Delete a defined name',
			requiredFields: ['name'],
			optionalFields: ['scope'],
		},
		{
			op: 'setStyle',
			description: 'Apply styles to a range',
			requiredFields: ['sheet', 'range', 'style'],
		},
		{
			op: 'freezePane',
			description: 'Freeze rows and columns',
			requiredFields: ['sheet', 'row', 'col'],
		},
		{ op: 'deleteComment', description: 'Delete a cell comment', requiredFields: ['sheet', 'ref'] },
		{
			op: 'deleteHyperlink',
			description: 'Delete a cell hyperlink',
			requiredFields: ['sheet', 'ref'],
		},
		{
			op: 'setDataValidation',
			description: 'Set data validation rule for a range',
			requiredFields: ['sheet', 'range', 'rule'],
		},
		{
			op: 'deleteDataValidation',
			description: 'Remove data validation from a range',
			requiredFields: ['sheet', 'range'],
		},
		{
			op: 'setAutoFilter',
			description: 'Enable auto-filter on a range',
			requiredFields: ['sheet', 'range'],
		},
		{
			op: 'clearAutoFilter',
			description: 'Clear auto-filter from a sheet',
			requiredFields: ['sheet'],
		},
		{
			op: 'setSheetProtection',
			description: 'Protect a sheet',
			requiredFields: ['sheet'],
			optionalFields: ['password', 'options'],
		},
		{ op: 'setTabColor', description: 'Set sheet tab color', requiredFields: ['sheet', 'color'] },
		{
			op: 'hideSheet',
			description: 'Hide or show a sheet',
			requiredFields: ['sheet'],
			optionalFields: ['hidden'],
		},
		{
			op: 'hideRows',
			description: 'Hide or show rows',
			requiredFields: ['sheet', 'at', 'count'],
			optionalFields: ['hidden'],
		},
		{
			op: 'hideCols',
			description: 'Hide or show columns',
			requiredFields: ['sheet', 'at', 'count'],
			optionalFields: ['hidden'],
		},
		{
			op: 'copySheet',
			description: 'Copy a sheet',
			requiredFields: ['sheet', 'newName'],
			optionalFields: ['position'],
		},
		{
			op: 'setConditionalFormat',
			description: 'Set conditional formatting rule',
			requiredFields: ['sheet', 'range', 'rule'],
		},
		{
			op: 'deleteConditionalFormat',
			description: 'Remove conditional formatting from a range',
			requiredFields: ['sheet', 'range'],
		},
		{
			op: 'setPageSetup',
			description: 'Set page setup for printing',
			requiredFields: ['sheet', 'setup'],
		},
		{ op: 'setPrintArea', description: 'Set print area', requiredFields: ['sheet', 'range'] },
		{
			op: 'copyRange',
			description: 'Copy a range to another location',
			requiredFields: ['sheet', 'source', 'target'],
		},
		{
			op: 'moveRange',
			description: 'Move a range to another location',
			requiredFields: ['sheet', 'source', 'target'],
		},
		{
			op: 'groupRows',
			description: 'Group rows for outlining',
			requiredFields: ['sheet', 'from', 'to'],
			optionalFields: ['collapsed', 'summaryBelow'],
		},
		{
			op: 'groupCols',
			description: 'Group columns for outlining',
			requiredFields: ['sheet', 'from', 'to'],
			optionalFields: ['collapsed', 'summaryRight'],
		},
		{
			op: 'setRichText',
			description: 'Set rich text in a cell',
			requiredFields: ['sheet', 'ref', 'runs'],
		},
	]
}

export function setCell(sheet: string, ref: string, value: InputValue): Operation {
	return { op: 'setCells', sheet, updates: [{ ref, value }] }
}

export function setFormula(sheet: string, ref: string, formula: string): Operation {
	return { op: 'setFormula', sheet, ref, formula }
}

export function addSheet(name: string): Operation {
	return { op: 'addSheet', name }
}

export function deleteSheet(sheet: string): Operation {
	return { op: 'deleteSheet', sheet }
}

export function renameSheet(sheet: string, newName: string): Operation {
	return { op: 'renameSheet', sheet, newName }
}

export function insertRows(sheet: string, at: number, count: number): Operation {
	return { op: 'insertRows', sheet, at, count }
}

export function deleteRows(sheet: string, at: number, count: number): Operation {
	return { op: 'deleteRows', sheet, at, count }
}

export function insertCols(sheet: string, at: number, count: number): Operation {
	return { op: 'insertCols', sheet, at, count }
}

export function deleteCols(sheet: string, at: number, count: number): Operation {
	return { op: 'deleteCols', sheet, at, count }
}
