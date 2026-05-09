import type { InputValue, Operation } from '@ascend/schema'

export interface OperationSchema {
	readonly op: string
	readonly description: string
	readonly requiredFields: readonly string[]
	readonly optionalFields?: readonly string[]
	readonly recoveryActions: readonly string[]
	readonly approval?: OperationApprovalMetadata
}

export interface OperationJsonSchema {
	readonly op: string
	readonly description: string
	readonly schemaDialect: 'json-schema-draft-2020-12-compatible'
	readonly standardSchema: {
		readonly version: 1
		readonly vendor: 'ascend'
		readonly name: string
	}
	readonly schema: {
		readonly type: 'object'
		readonly required: readonly string[]
		readonly properties: Record<
			string,
			{ readonly type: string; readonly description?: string; readonly enum?: readonly string[] }
		>
	}
	readonly examples: readonly Record<string, unknown>[]
	readonly invalidExamples: readonly OperationInvalidExample[]
	readonly recoveryActions: readonly string[]
	readonly approval?: OperationApprovalMetadata
}

export type ParseOperationsResult =
	| { readonly ok: true; readonly value: readonly Operation[] }
	| { readonly ok: false; readonly error: string; readonly issues: readonly string[] }

export interface OperationApprovalMetadata {
	readonly required: boolean
	readonly reason: string
	readonly approvalHint: string
}

export interface OperationInvalidExample {
	readonly input: Record<string, unknown>
	readonly issue: string
	readonly recoveryAction: string
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
	style: {
		type: 'object',
		description:
			'Style object with nested font, fill, border, alignment, numberFormat, protection. ' +
			'font: { name?, size?, bold?, italic?, underline?, strikethrough?, color?: { kind, rgb/theme/index } }. ' +
			'fill: { pattern?, fgColor?, bgColor? }. ' +
			'border: { top/bottom/left/right/diagonal?: { style?, color? }, diagonalUp?, diagonalDown? }. ' +
			'alignment: { horizontal?, vertical?, wrapText?, shrinkToFit?, textRotation?, indent? }. ' +
			'numberFormat: string (e.g. "#,##0.00"). ' +
			'protection: { locked?, hidden? }.',
	},
	password: { type: 'string', description: 'Protection password' },
	options: {
		type: 'object',
		description:
			'Sheet protection options: formatCells?, formatColumns?, formatRows?, insertColumns?, insertRows?, deleteColumns?, deleteRows?, sort?, autoFilter? (all boolean).',
	},
	protection: {
		type: 'object',
		description:
			'Workbook protection: lockStructure?, lockWindows?, lockRevision?, workbookPassword?, revisionsPassword?, plus optional Excel hash fields (workbookAlgorithmName, workbookHashValue, workbookSaltValue, workbookSpinCount, revisionsAlgorithmName, revisionsHashValue, revisionsSaltValue, revisionsSpinCount).',
	},
	color: { type: 'string', description: 'Color (hex or theme)' },
	hidden: { type: 'boolean', description: 'Whether to hide' },
	rule: {
		type: 'object',
		description:
			'For setDataValidation: { type: "list"|"whole"|"decimal"|"date"|"time"|"textLength"|"custom", formula1?, formula2?, operator?, allowBlank?, showErrorMessage?, errorTitle?, errorMessage?, showInputMessage?, promptTitle?, prompt? }. ' +
			'For setConditionalFormat: { type: "cellIs"|"expression"|"colorScale"|"dataBar"|"iconSet"|"top10"|"aboveAverage"|"duplicateValues"|"containsText", operator?, formula?, formula2?, priority?, stopIfTrue?, style?, colorScale?, dataBar?, iconSet? }.',
	},
	setup: {
		type: 'object',
		description:
			'Page setup: { orientation?: "portrait"|"landscape", paperSize?: number, scale?, fitToWidth?, fitToHeight?, margins?: { left?, right?, top?, bottom?, header?, footer? } }.',
	},
	source: { type: 'string', description: 'Source range' },
	target: { type: 'string', description: 'Target range' },
	newTarget: { type: 'string', description: 'Replacement external workbook path or URL' },
	targetPath: {
		type: 'string',
		description: 'XLSX package part path, such as xl/media/image1.png',
	},
	relId: { type: 'string', description: 'Relationship id inside the drawing part' },
	linkRelId: { type: 'string', description: 'Relationship id inside the external link part' },
	targetMode: { type: 'string', description: 'Relationship target mode, usually External' },
	contentBase64: { type: 'string', description: 'Base64-encoded binary content' },
	contentType: { type: 'string', description: 'MIME content type, such as image/png' },
	imageIndex: { type: 'integer', description: 'Zero-based image index on the sheet' },
	drawingPartPath: {
		type: 'string',
		description: 'Drawing XML part path, such as xl/drawings/drawing1.xml',
	},
	description: { type: 'string', description: 'Accessible image description' },
	anchor: { type: 'object', description: 'Image anchor object: oneCell, twoCell, or absolute' },
	chartIndex: { type: 'integer', description: 'Zero-based chart index after sheet/part filtering' },
	seriesIndex: { type: 'integer', description: 'Zero-based chart series index' },
	nameRef: { type: 'string', description: 'A1 formula reference for the chart series name' },
	categoryRef: { type: 'string', description: 'A1 formula reference for category or x values' },
	valueRef: { type: 'string', description: 'A1 formula reference for numeric values' },
	cacheId: { type: 'integer', description: 'Pivot cache id' },
	partPath: { type: 'string', description: 'XLSX package part path' },
	pivotTable: { type: 'string', description: 'Pivot table name that uses the cache' },
	sourceSheet: { type: 'string', description: 'Worksheet name for a pivot cache source' },
	sourceRef: { type: 'string', description: 'A1 range for a pivot cache source' },
	refreshOnLoad: { type: 'boolean', description: 'Whether Excel should refresh the cache on open' },
	enableRefresh: { type: 'boolean', description: 'Whether refresh is enabled for the cache' },
	invalid: { type: 'boolean', description: 'Whether the cache should be treated as stale' },
	saveData: { type: 'boolean', description: 'Whether cache records are saved in the workbook' },
	from: { type: 'integer', description: 'Start row/col index' },
	to: { type: 'integer', description: 'End row/col index' },
	collapsed: { type: 'boolean', description: 'Whether group is collapsed' },
	summaryBelow: { type: 'boolean', description: 'Summary row below' },
	summaryRight: { type: 'boolean', description: 'Summary column to right' },
	runs: { type: 'array', description: 'Rich text runs' },
}

export function listOperations(): readonly OperationSchema[] {
	return enrichOperationSchemas([
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
		{
			op: 'setWorkbookProtection',
			description: 'Set workbook-level protection metadata',
			requiredFields: ['protection'],
		},
		{ op: 'deleteTable', description: 'Remove a table from its sheet', requiredFields: ['table'] },
		{
			op: 'renameTable',
			description: 'Rename an existing table',
			requiredFields: ['table', 'newName'],
		},
		{
			op: 'resizeTable',
			description: 'Change a table range (rebuilds columns if width changes)',
			requiredFields: ['table', 'ref'],
		},
		{
			op: 'replaceImage',
			description: 'Replace image media bytes while preserving the existing drawing anchor',
			requiredFields: ['sheet', 'contentBase64', 'contentType'],
			optionalFields: ['targetPath', 'relId', 'name', 'imageIndex'],
		},
		{
			op: 'insertImage',
			description: 'Insert a new image with generated media and drawing relationships',
			requiredFields: ['sheet', 'contentBase64', 'contentType'],
			optionalFields: ['targetPath', 'drawingPartPath', 'relId', 'name', 'description', 'anchor'],
		},
		{
			op: 'deleteImage',
			description: 'Delete an image selected by targetPath, relId, name, or imageIndex',
			requiredFields: ['sheet'],
			optionalFields: ['targetPath', 'relId', 'name', 'imageIndex'],
		},
		{
			op: 'setChartSeriesSource',
			description:
				'Edit a chart series source references while preserving unsupported chart styling',
			requiredFields: ['seriesIndex'],
			optionalFields: ['partPath', 'sheet', 'chartIndex', 'nameRef', 'categoryRef', 'valueRef'],
		},
		{
			op: 'setPivotCache',
			description: 'Edit pivot cache source and refresh metadata without recalculating output',
			requiredFields: [],
			optionalFields: [
				'cacheId',
				'partPath',
				'pivotTable',
				'sourceSheet',
				'sourceRef',
				'refreshOnLoad',
				'enableRefresh',
				'invalid',
				'saveData',
			],
		},
		{
			op: 'rewriteExternalLink',
			description: 'Rewrite an external workbook link target while preserving link package parts',
			requiredFields: ['newTarget'],
			optionalFields: ['partPath', 'relId', 'linkRelId', 'target', 'targetMode'],
		},
	])
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
			schemaDialect: 'json-schema-draft-2020-12-compatible',
			standardSchema: { version: 1, vendor: 'ascend', name: op.op },
			schema: {
				type: 'object',
				required,
				properties,
			},
			examples: [operationExample(op.op)],
			invalidExamples: invalidOperationExamples(op),
			recoveryActions: op.recoveryActions,
			...(op.approval ? { approval: op.approval } : {}),
		}
	})
}

export function parseOperations(input: unknown): ParseOperationsResult {
	if (!Array.isArray(input)) {
		return {
			ok: false,
			error: 'Operations payload must be an array',
			issues: ['ops must be an array'],
		}
	}
	const schemas = new Map(listOperations().map((op) => [op.op, op]))
	const issues: string[] = []
	const ops: Operation[] = []
	input.forEach((item, index) => {
		if (item === null || typeof item !== 'object' || Array.isArray(item)) {
			issues.push(`ops[${index}] must be an object`)
			return
		}
		const record = item as Record<string, unknown>
		const op = record.op
		if (typeof op !== 'string') {
			issues.push(`ops[${index}].op must be a string`)
			return
		}
		const schema = schemas.get(op)
		if (!schema) {
			issues.push(`ops[${index}].op "${op}" is not supported`)
			return
		}
		for (const field of schema.requiredFields) {
			if (!(field in record)) issues.push(`ops[${index}].${field} is required for ${op}`)
		}
		ops.push(record as Operation)
	})
	if (issues.length > 0) return { ok: false, error: issues[0] ?? 'Invalid operations', issues }
	return { ok: true, value: ops }
}

function enrichOperationSchemas(
	ops: readonly Omit<OperationSchema, 'recoveryActions' | 'approval'>[],
): readonly OperationSchema[] {
	return ops.map((op) => {
		const approval = operationApproval(op.op)
		return {
			...op,
			recoveryActions: operationRecoveryActions(op.op),
			...(approval ? { approval } : {}),
		}
	})
}

function operationApproval(op: string): OperationApprovalMetadata | undefined {
	if (
		op === 'deleteSheet' ||
		op === 'deleteRows' ||
		op === 'deleteCols' ||
		op === 'deleteTable' ||
		op === 'deleteDefinedName'
	) {
		return {
			required: true,
			reason: 'This operation can remove workbook content or metadata.',
			approvalHint: 'Run ascend plan first and pass the emitted approval id to ascend commit.',
		}
	}
	if (op === 'clearRange') {
		return {
			required: true,
			reason:
				'Approval is required when what is "all" because values, formulas, and styles are removed.',
			approvalHint: 'Prefer what="values" or what="formulas" when full clearing is not intended.',
		}
	}
	return undefined
}

function operationRecoveryActions(op: string): readonly string[] {
	const common = [
		'Run ascend ops --op <operation> --json for the canonical schema and examples.',
		'Run ascend plan <file> --ops ops.json --json before commit.',
	]
	switch (op) {
		case 'setCells':
			return ['Ensure updates is a non-empty array of { ref, value } objects.', ...common]
		case 'setFormula':
		case 'fillFormula':
			return [
				'Use formula text with or without a leading "=" and verify with ascend plan.',
				...common,
			]
		case 'deleteSheet':
		case 'deleteRows':
		case 'deleteCols':
		case 'deleteTable':
		case 'deleteDefinedName':
			return [
				'Review plan.approvals and provide --approval only for intentional deletion.',
				...common,
			]
		case 'clearRange':
			return [
				'Use what="values", "formulas", or "styles" unless all content should be removed.',
				...common,
			]
		case 'createTable':
		case 'resizeTable':
			return ['Confirm the table range includes the intended header and data rows.', ...common]
		case 'replaceImage':
			return [
				'Use inspect --detail images or visualInventory to select targetPath, relId, name, or imageIndex.',
				'Ensure contentBase64 bytes match contentType.',
				...common,
			]
		case 'insertImage':
			return [
				'Provide image/* contentType and base64 image bytes; omit targetPath/relId for automatic allocation.',
				'Use an anchor object to position the inserted image.',
				...common,
			]
		case 'deleteImage':
			return [
				'Use inspect --detail images or visualInventory to select targetPath, relId, name, or imageIndex.',
				'Run plan before commit to confirm the selected image is the only match.',
				...common,
			]
		case 'setChartSeriesSource':
			return [
				'Use inspect --detail visuals or visualInventory to select partPath, sheet, chartIndex, and seriesIndex.',
				'Provide at least one of nameRef, categoryRef, or valueRef.',
				...common,
			]
		case 'setPivotCache':
			return [
				'Use inspect --detail pivots to choose cacheId, partPath, or pivotTable.',
				'Set invalid=true and refreshOnLoad=true when changing source ranges without recalculating pivot output.',
				...common,
			]
		case 'rewriteExternalLink':
			return [
				'Use inspect --detail external-refs to choose partPath, relId, linkRelId, or target.',
				'Run ascend plan before commit to verify package preservation and formula-state warnings.',
				...common,
			]
		default:
			return common
	}
}

function invalidOperationExamples(op: OperationSchema): readonly OperationInvalidExample[] {
	const missingField = op.requiredFields[0]
	return [
		{
			input: { op: op.op },
			issue: missingField
				? `Missing required field "${missingField}".`
				: 'Operation is missing required fields.',
			recoveryAction: `Add all required fields: ${['op', ...op.requiredFields].join(', ')}.`,
		},
	]
}

function operationExample(op: string): Record<string, unknown> {
	switch (op) {
		case 'setCells':
			return { op, sheet: 'Sheet1', updates: [{ ref: 'A1', value: 123 }] }
		case 'setFormula':
			return { op, sheet: 'Sheet1', ref: 'B1', formula: '=SUM(A1:A10)' }
		case 'fillFormula':
			return { op, sheet: 'Sheet1', range: 'B2:B10', formula: '=A2*2' }
		case 'clearRange':
			return { op, sheet: 'Sheet1', range: 'A1:B10', what: 'values' }
		case 'insertRows':
		case 'deleteRows':
		case 'hideRows':
			return { op, sheet: 'Sheet1', at: 1, count: 2 }
		case 'insertCols':
		case 'deleteCols':
		case 'hideCols':
			return { op, sheet: 'Sheet1', at: 1, count: 2 }
		case 'addSheet':
			return { op, name: 'Summary' }
		case 'deleteSheet':
			return { op, sheet: 'OldData' }
		case 'renameSheet':
			return { op, sheet: 'Sheet1', newName: 'Data' }
		case 'moveSheet':
			return { op, sheet: 'Data', position: 0 }
		case 'createTable':
			return { op, sheet: 'Sheet1', ref: 'A1:D20', name: 'Sales', hasHeaders: true }
		case 'appendRows':
			return { op, table: 'Sales', rows: [['West', 1200]] }
		case 'sortRange':
			return { op, sheet: 'Sheet1', range: 'A1:D20', by: [{ column: 1, descending: true }] }
		case 'mergeCells':
		case 'unmergeCells':
			return { op, sheet: 'Sheet1', range: 'A1:C1' }
		case 'setColWidth':
			return { op, sheet: 'Sheet1', col: 0, width: 18 }
		case 'setRowHeight':
			return { op, sheet: 'Sheet1', row: 0, height: 24 }
		case 'setComment':
			return { op, sheet: 'Sheet1', ref: 'A1', text: 'Reviewed', author: 'Ascend' }
		case 'setHyperlink':
			return { op, sheet: 'Sheet1', ref: 'A1', url: 'https://example.com', display: 'Example' }
		case 'setNumberFormat':
			return { op, sheet: 'Sheet1', range: 'A:A', format: '$#,##0.00' }
		case 'setDefinedName':
			return { op, name: 'SalesTotal', ref: 'Sheet1!$B$2:$B$20' }
		case 'deleteDefinedName':
			return { op, name: 'SalesTotal' }
		case 'setStyle':
			return { op, sheet: 'Sheet1', range: 'A1:D1', style: { font: { bold: true } } }
		case 'freezePane':
			return { op, sheet: 'Sheet1', row: 1, col: 0 }
		case 'deleteComment':
		case 'deleteHyperlink':
			return { op, sheet: 'Sheet1', ref: 'A1' }
		case 'setDataValidation':
			return { op, sheet: 'Sheet1', range: 'A2:A20', rule: { type: 'list', formula1: '"Yes,No"' } }
		case 'deleteDataValidation':
		case 'setAutoFilter':
		case 'deleteConditionalFormat':
		case 'setPrintArea':
			return { op, sheet: 'Sheet1', range: 'A1:D20' }
		case 'clearAutoFilter':
			return { op, sheet: 'Sheet1' }
		case 'setSheetProtection':
			return { op, sheet: 'Sheet1', options: { formatCells: false } }
		case 'setTabColor':
			return { op, sheet: 'Sheet1', color: '#4472C4' }
		case 'hideSheet':
			return { op, sheet: 'Sheet1', hidden: true }
		case 'copySheet':
			return { op, sheet: 'Sheet1', newName: 'Sheet1 Copy' }
		case 'setConditionalFormat':
			return {
				op,
				sheet: 'Sheet1',
				range: 'A2:A20',
				rule: { type: 'cellIs', operator: 'greaterThan', formula: '100' },
			}
		case 'setPageSetup':
			return { op, sheet: 'Sheet1', setup: { orientation: 'landscape' } }
		case 'copyRange':
		case 'moveRange':
			return { op, sheet: 'Sheet1', source: 'A1:B5', target: 'D1:E5' }
		case 'groupRows':
			return { op, sheet: 'Sheet1', from: 2, to: 10, collapsed: false }
		case 'groupCols':
			return { op, sheet: 'Sheet1', from: 2, to: 5, collapsed: false }
		case 'setRichText':
			return { op, sheet: 'Sheet1', ref: 'A1', runs: [{ text: 'Bold', font: { bold: true } }] }
		case 'setWorkbookProtection':
			return { op, protection: { lockStructure: true } }
		case 'deleteTable':
			return { op, table: 'Sales' }
		case 'renameTable':
			return { op, table: 'Sales', newName: 'SalesData' }
		case 'resizeTable':
			return { op, table: 'Sales', ref: 'A1:E20' }
		case 'replaceImage':
			return {
				op,
				sheet: 'Sheet1',
				targetPath: 'xl/media/image1.png',
				contentBase64: 'iVBORw0KGgo=',
				contentType: 'image/png',
			}
		case 'insertImage':
			return {
				op,
				sheet: 'Sheet1',
				contentBase64: 'iVBORw0KGgo=',
				contentType: 'image/png',
				name: 'Logo',
				anchor: { kind: 'oneCell', from: { row: 1, col: 1 }, cx: 320000, cy: 240000 },
			}
		case 'deleteImage':
			return { op, sheet: 'Sheet1', imageIndex: 0 }
		case 'setChartSeriesSource':
			return {
				op,
				partPath: 'xl/charts/chart1.xml',
				seriesIndex: 0,
				categoryRef: 'Data!$A$2:$A$20',
				valueRef: 'Data!$B$2:$B$20',
			}
		case 'setPivotCache':
			return {
				op,
				pivotTable: 'PivotTable1',
				sourceSheet: 'RawData',
				sourceRef: 'A1:E200',
				refreshOnLoad: true,
				invalid: true,
			}
		case 'rewriteExternalLink':
			return {
				op,
				partPath: 'xl/externalLinks/externalLink1.xml',
				linkRelId: 'rIdExt',
				newTarget: '../sources/reforecast.xlsx',
				targetMode: 'External',
			}
		default:
			return { op }
	}
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
