import type { InputValue, Operation } from '@ascend/schema'

export interface OperationSchema {
	readonly op: string
	readonly description: string
	readonly requiredFields: readonly string[]
	readonly optionalFields?: readonly string[]
}

export interface OperationJsonSchema {
	readonly op: string
	readonly description: string
	readonly schema: {
		readonly type: 'object'
		readonly required: readonly string[]
		readonly properties: Record<
			string,
			{ readonly type: string; readonly description?: string; readonly enum?: readonly string[] }
		>
	}
}

const FIELD_SCHEMAS: Record<
	string,
	{ type: string; description?: string; enum?: readonly string[] }
> = {
	sheet: { type: 'string', description: 'Sheet name' },
	updates: {
		type: 'array',
		description: 'Array of { ref: string, value: number|string|boolean|null }',
	},
	ref: { type: 'string', description: 'Cell reference (e.g. A1)' },
	formula: { type: 'string', description: 'Formula text' },
	range: { type: 'string', description: 'Cell range (e.g. A1:B10)' },
	what: {
		type: 'string',
		enum: ['values', 'formulas', 'styles', 'all'],
		description: 'What to clear',
	},
	at: { type: 'integer', description: 'Row or column index (0-based)' },
	count: { type: 'integer', description: 'Number of rows/columns' },
	name: { type: 'string', description: 'Sheet or table name' },
	position: { type: 'integer', description: 'Position index' },
	newName: { type: 'string', description: 'New name' },
	hasHeaders: { type: 'boolean', description: 'Whether range has header row' },
	table: { type: 'string', description: 'Table name' },
	rows: { type: 'array', description: 'Array of row arrays' },
	by: { type: 'array', description: 'Sort specs: [{ column, descending? }]' },
	col: { type: 'integer', description: 'Column index' },
	width: { type: 'number', description: 'Column width' },
	row: { type: 'integer', description: 'Row index' },
	height: { type: 'number', description: 'Row height' },
	text: { type: 'string', description: 'Comment or cell text' },
	author: { type: 'string', description: 'Comment author' },
	url: { type: 'string', description: 'Hyperlink URL' },
	display: { type: 'string', description: 'Display text for hyperlink' },
	format: { type: 'string', description: 'Number format code' },
	scope: { type: 'string', description: 'Scope (workbook or sheet name)' },
	style: { type: 'object', description: 'Style object (font, fill, border, alignment)' },
	password: { type: 'string', description: 'Protection password' },
	options: { type: 'object', description: 'Protection options' },
	color: { type: 'string', description: 'Color (hex or theme)' },
	hidden: { type: 'boolean', description: 'Whether to hide' },
	rule: { type: 'object', description: 'Validation or format rule' },
	setup: { type: 'object', description: 'Page setup object' },
	source: { type: 'string', description: 'Source range' },
	target: { type: 'string', description: 'Target range' },
	from: { type: 'integer', description: 'Start row/col index' },
	to: { type: 'integer', description: 'End row/col index' },
	collapsed: { type: 'boolean', description: 'Whether group is collapsed' },
	summaryBelow: { type: 'boolean', description: 'Summary row below' },
	summaryRight: { type: 'boolean', description: 'Summary column to right' },
	runs: { type: 'array', description: 'Rich text runs' },
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

export function getOperationsSchema(): readonly OperationJsonSchema[] {
	const ops = listOperations()
	return ops.map((op) => {
		const required = ['op', ...op.requiredFields] as const
		const properties: Record<
			string,
			{ type: string; description?: string; enum?: readonly string[] }
		> = {
			op: {
				type: 'string',
				enum: [op.op],
				description: 'Operation type',
			},
		}
		for (const field of [...op.requiredFields, ...(op.optionalFields ?? [])]) {
			const schema = FIELD_SCHEMAS[field] ?? { type: 'string' }
			properties[field] = schema
		}
		return {
			op: op.op,
			description: op.description,
			schema: {
				type: 'object',
				required,
				properties,
			},
		}
	})
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
