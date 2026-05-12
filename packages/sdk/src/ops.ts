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
			{
				readonly type: string | readonly string[]
				readonly description?: string
				readonly enum?: readonly string[]
			}
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
	{ type: string | readonly string[]; description?: string; enum?: readonly string[] }
> = {
	sheet: { type: 'string', description: 'Sheet name' },
	updates: {
		type: 'array',
		description: 'Array of { ref: string, value: number|string|boolean|null }',
	},
	ref: { type: 'string', description: 'Cell reference (e.g. A1)' },
	formula: { type: 'string', description: 'Formula text' },
	range: { type: 'string', description: 'Cell or sparkline source range (e.g. A1:B10)' },
	what: {
		type: 'string',
		enum: ['values', 'formulas', 'styles', 'all'],
		description: 'What to clear',
	},
	mode: {
		type: 'string',
		enum: [
			'merge',
			'replace',
			'all',
			'values',
			'formulas',
			'formats',
			'styles',
			'validations',
			'comments',
			'hyperlinks',
			'append',
		],
		description:
			'Paste mode for copy/move, workbook property mode (merge|replace), or conditional-format update mode (replace|append).',
	},
	priority: { type: 'integer', description: 'Conditional-format rule priority' },
	ruleIndex: { type: 'integer', description: 'Zero-based conditional-format rule index' },
	reassignPriorities: {
		type: 'boolean',
		description: 'Whether to rewrite conditional-format priorities to match stored rule order',
	},
	index: { type: 'integer', description: 'Workbook view index' },
	at: { type: 'integer', description: 'Row or column index (0-based)' },
	count: { type: 'integer', description: 'Number of rows/columns' },
	name: {
		type: 'string',
		description: 'Sheet or table name; table names must be workbook-unique case-insensitively',
	},
	position: { type: 'integer', description: 'Position index' },
	newName: { type: 'string', description: 'New name' },
	type: { type: 'string', description: 'Sparkline type such as line, column, or stacked' },
	hasHeaders: { type: 'boolean', description: 'Whether range has header row' },
	table: {
		type: 'string',
		description: 'Workbook-unique table name; run check if imported table names may be ambiguous',
	},
	rows: { type: 'array', description: 'Array of row arrays' },
	column: { type: ['string', 'integer'], description: 'Table column name or 0-based column index' },
	totalsRowFunction: { type: 'string', description: 'Excel table totalsRowFunction value' },
	totalsRowFormula: { type: 'string', description: 'Formula for the table totals row cell' },
	totalsRowLabel: { type: 'string', description: 'Label for the table totals row cell' },
	styleName: {
		type: ['string', 'null'],
		description: 'Excel table style name; use null to clear the style name',
	},
	showFirstColumn: {
		type: 'boolean',
		description: 'Whether the table style highlights first column',
	},
	showLastColumn: {
		type: 'boolean',
		description: 'Whether the table style highlights last column',
	},
	showRowStripes: { type: 'boolean', description: 'Whether the table style shows row stripes' },
	showColumnStripes: {
		type: 'boolean',
		description: 'Whether the table style shows column stripes',
	},
	by: { type: 'array', description: 'Sort specs: [{ column, descending? }]' },
	col: { type: 'integer', description: 'Column index' },
	width: { type: 'number', description: 'Column width' },
	row: { type: 'integer', description: 'Row index' },
	height: { type: 'number', description: 'Row height' },
	text: { type: 'string', description: 'Comment or cell text' },
	author: { type: 'string', description: 'Comment author' },
	url: { type: 'string', description: 'External hyperlink URL' },
	location: { type: 'string', description: 'Internal workbook hyperlink location, e.g. Sheet2!A1' },
	display: { type: 'string', description: 'Display text for hyperlink' },
	tooltip: { type: 'string', description: 'Tooltip text for hyperlink' },
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
	properties: {
		type: 'object',
		description:
			'Workbook properties: codeName?, defaultThemeVersion?, filterPrivacy?, date1904?. Use null values to clear individual properties.',
	},
	view: {
		type: 'object',
		description:
			'Workbook view metadata: activeTab?, firstSheet?, visibility?, tabRatio?. Use null values to clear individual fields, or null view to delete a view.',
	},
	settings: {
		type: 'object',
		description:
			'Calculation settings: calcMode?, fullCalcOnLoad?, calcCompleted?, calcOnSave?, forceFullCalc?, calcId?, dateSystem?, iterativeCalc?.',
	},
	themeName: { type: 'string', description: 'Workbook theme name' },
	colorSchemeName: {
		type: 'string',
		description: 'Theme color scheme name',
	},
	majorFontLatin: {
		type: 'string',
		description: 'Theme major Latin font typeface',
	},
	minorFontLatin: {
		type: 'string',
		description: 'Theme minor Latin font typeface',
	},
	themeColors: {
		type: 'array',
		description:
			'Theme color updates: [{ slot, rgb? }] or [{ slot, systemColor, lastColor? }]. Slots: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink.',
	},
	color: { type: 'string', description: 'Color (hex or theme)' },
	hidden: { type: 'boolean', description: 'Whether to hide' },
	rule: {
		type: 'object',
		description:
			'For setDataValidation: { type: "list"|"whole"|"decimal"|"date"|"time"|"textLength"|"custom", formula1?, formula2?, operator?, allowBlank?, showDropDown?, showErrorMessage?, errorTitle?, errorMessage?, errorStyle?, imeMode?, showInputMessage?, promptTitle?, prompt? }. ' +
			'For setConditionalFormat: { type: "cellIs"|"expression"|"colorScale"|"dataBar"|"iconSet"|"top10"|"aboveAverage"|"duplicateValues"|"containsText", operator?, formula?, formula2?, priority?, stopIfTrue?, style?, colorScale?, dataBar?, iconSet? }. Use mode="append" to preserve existing overlapping rules.',
	},
	setup: {
		type: 'object',
		description:
			'Page setup: { orientation?: "portrait"|"landscape", paperSize?: number, scale?, fitToWidth?, fitToHeight?, margins?: { left?, right?, top?, bottom?, header?, footer? } }.',
	},
	source: { type: 'string', description: 'Source range' },
	target: { type: 'string', description: 'Target range' },
	targetSheet: {
		type: 'string',
		description: 'Destination sheet for copyRange or moveRange; defaults to the source sheet',
	},
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
	drawingObjectIndex: {
		type: 'integer',
		description: 'Zero-based drawing object index on the sheet',
	},
	groupIndex: { type: 'integer', description: 'Zero-based sparkline group index on the sheet' },
	filterIndex: {
		type: 'integer',
		description: 'Zero-based custom sheet view advanced filter index on the sheet',
	},
	values: { type: 'array', description: 'Filter value-list strings for the selected column' },
	sortRef: { type: 'string', description: 'Advanced filter sort range' },
	sortBy: { type: 'string', description: 'Advanced filter sort condition range' },
	locationRange: { type: 'string', description: 'Sparkline destination sqref range' },
	markers: { type: 'boolean', description: 'Whether sparkline markers are displayed' },
	highPoint: { type: 'boolean', description: 'Whether sparkline high points are highlighted' },
	lowPoint: { type: 'boolean', description: 'Whether sparkline low points are highlighted' },
	firstPoint: { type: 'boolean', description: 'Whether sparkline first points are highlighted' },
	lastPoint: { type: 'boolean', description: 'Whether sparkline last points are highlighted' },
	negative: { type: 'boolean', description: 'Whether sparkline negative points are highlighted' },
	displayXAxis: { type: 'boolean', description: 'Whether sparkline x-axis is displayed' },
	id: { type: 'integer', description: 'Drawing object non-visual id' },
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
	connectionId: { type: 'integer', description: 'Workbook connection id' },
	partPath: { type: 'string', description: 'XLSX package part path' },
	threadedCommentId: { type: 'string', description: 'Threaded comment id attribute' },
	commentIndex: { type: 'integer', description: 'Zero-based threaded comment index on the sheet' },
	pivotTable: { type: 'string', description: 'Pivot table name that uses the cache' },
	fieldIndex: { type: 'integer', description: 'Zero-based pivot field index' },
	itemIndex: { type: 'integer', description: 'Zero-based pivot field item index' },
	showDetails: {
		type: ['boolean', 'null'],
		description: 'Pivot field item show-details state; use null to clear',
	},
	manualFilter: {
		type: ['boolean', 'null'],
		description: 'Pivot field item manual-filter state; use null to clear',
	},
	selectedPageItem: {
		type: ['integer', 'null'],
		description: 'Selected page-field item index; use null to clear',
	},
	slicerCache: { type: 'string', description: 'Slicer cache name or package part path' },
	timelineCache: { type: 'string', description: 'Timeline cache name or package part path' },
	startDate: { type: 'string', description: 'Timeline selection start date-time string' },
	endDate: { type: 'string', description: 'Timeline selection end date-time string' },
	item: { type: 'integer', description: 'Zero-based slicer cache item x index' },
	selected: { type: 'boolean', description: 'Slicer item selected state; use null to clear' },
	noData: { type: 'boolean', description: 'Slicer item no-data state; use null to clear' },
	sourceSheet: { type: 'string', description: 'Worksheet name for a pivot cache source' },
	sourceRef: { type: 'string', description: 'A1 range for a pivot cache source' },
	refreshOnLoad: { type: 'boolean', description: 'Whether Excel should refresh the cache on open' },
	enableRefresh: { type: 'boolean', description: 'Whether refresh is enabled for the cache' },
	invalid: { type: 'boolean', description: 'Whether the cache should be treated as stale' },
	saveData: { type: 'boolean', description: 'Whether cache records are saved in the workbook' },
	refreshedVersion: { type: 'integer', description: 'Excel refresh engine version metadata' },
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
		{
			op: 'insertRows',
			description: 'Insert rows only when shifted table ranges remain non-overlapping',
			requiredFields: ['sheet', 'at', 'count'],
		},
		{
			op: 'deleteRows',
			description:
				'Delete rows without partially removing table header or totals rows unless deleting the full table span',
			requiredFields: ['sheet', 'at', 'count'],
		},
		{
			op: 'insertCols',
			description: 'Insert columns only when shifted table ranges remain non-overlapping',
			requiredFields: ['sheet', 'at', 'count'],
		},
		{
			op: 'deleteCols',
			description:
				'Delete columns only after structured references to removed table fields are rewritten or removed',
			requiredFields: ['sheet', 'at', 'count'],
		},
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
			description: 'Create a workbook-unique table from a non-overlapping range',
			requiredFields: ['sheet', 'ref', 'name', 'hasHeaders'],
		},
		{
			op: 'appendRows',
			description: 'Append rows to a table without expanding into or shifting another table range',
			requiredFields: ['table', 'rows'],
		},
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
			requiredFields: ['sheet', 'ref'],
			optionalFields: ['url', 'location', 'display', 'tooltip'],
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
			description: 'Enable or edit a worksheet auto-filter while preserving existing criteria',
			requiredFields: ['sheet', 'range'],
			optionalFields: ['column', 'values', 'sortRef', 'sortBy', 'descending'],
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
			description: 'Set or append a conditional formatting rule while preserving priority order',
			requiredFields: ['sheet', 'range', 'rule'],
			optionalFields: ['mode', 'reassignPriorities'],
		},
		{
			op: 'deleteConditionalFormat',
			description: 'Remove conditional formatting from a range or delete a specific rule',
			requiredFields: ['sheet'],
			optionalFields: ['range', 'priority', 'ruleIndex'],
		},
		{
			op: 'setPageSetup',
			description: 'Set page setup for printing',
			requiredFields: ['sheet', 'setup'],
		},
		{ op: 'setPrintArea', description: 'Set print area', requiredFields: ['sheet', 'range'] },
		{
			op: 'copyRange',
			description: 'Copy a range to another location with optional Excel-like paste mode',
			requiredFields: ['sheet', 'source', 'target'],
			optionalFields: ['targetSheet', 'mode'],
		},
		{
			op: 'moveRange',
			description: 'Move a range to another location with optional Excel-like paste mode',
			requiredFields: ['sheet', 'source', 'target'],
			optionalFields: ['targetSheet', 'mode'],
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
			op: 'setWorkbookProperties',
			description: 'Set workbook-level package properties such as codeName and date system',
			requiredFields: ['properties'],
			optionalFields: ['mode'],
		},
		{
			op: 'setDocumentProperties',
			description: 'Set core, app, and custom package document properties',
			requiredFields: ['properties'],
			optionalFields: ['mode'],
		},
		{
			op: 'setWorkbookView',
			description: 'Set, append, merge, replace, or delete workbook view metadata',
			requiredFields: ['view'],
			optionalFields: ['index', 'mode'],
		},
		{
			op: 'setCalcSettings',
			description: 'Set workbook calculation settings and date system metadata',
			requiredFields: ['settings'],
		},
		{
			op: 'setTheme',
			description: 'Edit workbook theme names, fonts, and color slots',
			requiredFields: [],
			optionalFields: [
				'themeName',
				'colorSchemeName',
				'majorFontLatin',
				'minorFontLatin',
				'themeColors',
			],
		},
		{
			op: 'setWorkbookProtection',
			description: 'Set workbook-level protection metadata',
			requiredFields: ['protection'],
		},
		{
			op: 'deleteTable',
			description:
				'Remove table metadata only after structured references are rewritten or removed',
			requiredFields: ['table'],
		},
		{
			op: 'renameTable',
			description: 'Rename an existing table to a workbook-unique name',
			requiredFields: ['table', 'newName'],
		},
		{
			op: 'resizeTable',
			description:
				'Change a table range without overlapping another table or dropping referenced fields',
			requiredFields: ['table', 'ref'],
		},
		{
			op: 'setTableColumn',
			description: 'Rename a table column or set calculated-column formula and totals metadata',
			requiredFields: ['table', 'column'],
			optionalFields: [
				'newName',
				'formula',
				'totalsRowFunction',
				'totalsRowFormula',
				'totalsRowLabel',
			],
		},
		{
			op: 'setTableStyle',
			description: 'Edit table style name and display flags',
			requiredFields: ['table'],
			optionalFields: [
				'styleName',
				'showFirstColumn',
				'showLastColumn',
				'showRowStripes',
				'showColumnStripes',
			],
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
			op: 'setDrawingText',
			description: 'Edit an existing shape or text box text run while preserving drawing XML',
			requiredFields: ['sheet', 'text'],
			optionalFields: ['drawingPartPath', 'id', 'name', 'drawingObjectIndex'],
		},
		{
			op: 'setThreadedComment',
			description: 'Edit existing threaded comment text while preserving thread metadata',
			requiredFields: ['sheet', 'text'],
			optionalFields: ['partPath', 'threadedCommentId', 'ref', 'commentIndex'],
		},
		{
			op: 'setChartSeriesSource',
			description: 'Edit chart series source references while preserving opaque chart styling',
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
			op: 'setPivotFieldItem',
			description: 'Edit pivot field item filter flags while preserving pivot table XML',
			requiredFields: ['fieldIndex', 'itemIndex'],
			optionalFields: [
				'pivotTable',
				'partPath',
				'sheet',
				'hidden',
				'showDetails',
				'manualFilter',
				'selectedPageItem',
			],
		},
		{
			op: 'setSlicerCacheItem',
			description:
				'Edit tabular slicer cache item selected/no-data flags without recalculating output',
			requiredFields: ['item'],
			optionalFields: ['slicerCache', 'partPath', 'selected', 'noData'],
		},
		{
			op: 'setTimelineRange',
			description: 'Edit a timeline cache selected date range without recalculating pivot output',
			requiredFields: ['startDate', 'endDate'],
			optionalFields: ['timelineCache', 'partPath'],
		},
		{
			op: 'setSparklineGroup',
			description: 'Edit a preserved sparkline group source range and display flags',
			requiredFields: ['sheet', 'groupIndex'],
			optionalFields: [
				'range',
				'locationRange',
				'type',
				'markers',
				'highPoint',
				'lowPoint',
				'firstPoint',
				'lastPoint',
				'negative',
				'displayXAxis',
			],
		},
		{
			op: 'setAdvancedFilter',
			description: 'Edit a preserved custom sheet view advanced filter and sort criteria',
			requiredFields: ['sheet', 'filterIndex'],
			optionalFields: ['range', 'column', 'values', 'sortRef', 'sortBy', 'descending'],
		},
		{
			op: 'setConnectionRefresh',
			description: 'Edit workbook connection and query-table refresh metadata',
			requiredFields: [],
			optionalFields: [
				'partPath',
				'name',
				'connectionId',
				'sheet',
				'refreshOnLoad',
				'saveData',
				'refreshedVersion',
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
			{ type: string | readonly string[]; description?: string; enum?: readonly string[] }
		> = {
			op: {
				type: 'string',
				enum: [op.op],
				description: 'Operation type',
			},
		}
		for (const field of [...op.requiredFields, ...(op.optionalFields ?? [])]) {
			const schema =
				op.op === 'setPivotFieldItem' && field === 'hidden'
					? { ...FIELD_SCHEMAS.hidden, type: ['boolean', 'null'] }
					: (FIELD_SCHEMAS[field] ?? { type: 'string' })
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
		const allowedFields = new Set([
			'op',
			...schema.requiredFields,
			...(schema.optionalFields ?? []),
		])
		for (const field of Object.keys(record)) {
			if (!allowedFields.has(field)) issues.push(`ops[${index}].${field} is not valid for ${op}`)
		}
		for (const field of schema.requiredFields) {
			if (!(field in record)) issues.push(`ops[${index}].${field} is required for ${op}`)
		}
		for (const field of allowedFields) {
			if (field === 'op' || !(field in record)) continue
			const issue = validateOperationField(record, field, `ops[${index}].${field}`)
			if (issue) issues.push(issue)
		}
		if (
			op === 'setHyperlink' &&
			!hasNonEmptyString(record.url) &&
			!hasNonEmptyString(record.location)
		) {
			issues.push(`ops[${index}].url or ops[${index}].location is required for setHyperlink`)
		}
		ops.push(record as Operation)
	})
	if (issues.length > 0) return { ok: false, error: issues[0] ?? 'Invalid operations', issues }
	return { ok: true, value: ops }
}

function hasNonEmptyString(value: unknown): boolean {
	return typeof value === 'string' && value.trim().length > 0
}

const PASTE_MODES = new Set([
	'all',
	'values',
	'formulas',
	'formats',
	'styles',
	'validations',
	'comments',
	'hyperlinks',
])

function validateOperationField(
	record: Record<string, unknown>,
	field: string,
	path: string,
): string | null {
	const value = record[field]
	switch (field) {
		case 'sheet':
		case 'ref':
		case 'range':
		case 'name':
		case 'newName':
		case 'table':
		case 'text':
		case 'author':
		case 'url':
		case 'location':
		case 'display':
		case 'tooltip':
		case 'format':
		case 'scope':
		case 'password':
		case 'color':
		case 'source':
		case 'target':
		case 'targetSheet':
		case 'newTarget':
		case 'targetPath':
		case 'relId':
		case 'linkRelId':
		case 'targetMode':
		case 'contentBase64':
		case 'contentType':
		case 'drawingPartPath':
		case 'description':
		case 'nameRef':
		case 'categoryRef':
		case 'valueRef':
		case 'partPath':
		case 'locationRange':
		case 'type':
		case 'threadedCommentId':
		case 'pivotTable':
		case 'sourceSheet':
		case 'sourceRef':
		case 'slicerCache':
		case 'timelineCache':
		case 'startDate':
		case 'endDate':
		case 'sortRef':
		case 'sortBy':
			return typeof value === 'string' ? null : `${path} must be a string`
		case 'formula':
			if (record.op === 'setTableColumn') {
				return value === null || typeof value === 'string'
					? null
					: `${path} must be a string or null`
			}
			return typeof value === 'string' ? null : `${path} must be a string`
		case 'themeName':
		case 'colorSchemeName':
		case 'majorFontLatin':
		case 'minorFontLatin':
			return typeof value === 'string' ? null : `${path} must be a string`
		case 'what':
			return value === 'values' || value === 'formulas' || value === 'styles' || value === 'all'
				? null
				: `${path} must be one of values, formulas, styles, all`
		case 'mode':
			return validateMode(record.op, value, path)
		case 'totalsRowFunction':
		case 'totalsRowFormula':
		case 'totalsRowLabel':
			return value === null || typeof value === 'string' ? null : `${path} must be a string or null`
		case 'styleName':
			return value === null || typeof value === 'string' ? null : `${path} must be a string or null`
		case 'at':
		case 'position':
		case 'col':
		case 'row':
		case 'ruleIndex':
		case 'imageIndex':
		case 'drawingObjectIndex':
		case 'groupIndex':
		case 'filterIndex':
		case 'commentIndex':
		case 'id':
		case 'chartIndex':
		case 'seriesIndex':
		case 'cacheId':
		case 'connectionId':
		case 'item':
		case 'fieldIndex':
		case 'itemIndex':
		case 'from':
		case 'to':
		case 'index':
		case 'refreshedVersion':
			return isNonNegativeInteger(value) ? null : `${path} must be a non-negative integer`
		case 'selectedPageItem':
			return value === null || isNonNegativeInteger(value)
				? null
				: `${path} must be a non-negative integer or null`
		case 'priority':
		case 'count':
			return isPositiveInteger(value) ? null : `${path} must be a positive integer`
		case 'width':
		case 'height':
			return isFiniteNumber(value) ? null : `${path} must be a finite number`
		case 'hasHeaders':
		case 'descending':
		case 'reassignPriorities':
		case 'showFirstColumn':
		case 'showLastColumn':
		case 'showRowStripes':
		case 'showColumnStripes':
		case 'collapsed':
		case 'summaryBelow':
		case 'summaryRight':
		case 'refreshOnLoad':
		case 'enableRefresh':
		case 'invalid':
		case 'saveData':
		case 'markers':
		case 'highPoint':
		case 'lowPoint':
		case 'firstPoint':
		case 'lastPoint':
		case 'negative':
		case 'displayXAxis':
			return typeof value === 'boolean' ? null : `${path} must be a boolean`
		case 'showDetails':
		case 'manualFilter':
			return value === null || typeof value === 'boolean'
				? null
				: `${path} must be a boolean or null`
		case 'hidden':
			if (record.op === 'setPivotFieldItem') {
				return value === null || typeof value === 'boolean'
					? null
					: `${path} must be a boolean or null`
			}
			return typeof value === 'boolean' ? null : `${path} must be a boolean`
		case 'selected':
		case 'noData':
			return value === null || typeof value === 'boolean'
				? null
				: `${path} must be a boolean or null`
		case 'column':
			return typeof value === 'string' || isNonNegativeInteger(value)
				? null
				: `${path} must be a string or non-negative integer`
		case 'updates':
			return validateUpdates(value, path)
		case 'rows':
			return validateRows(value, path)
		case 'values':
			return validateStringArray(value, path)
		case 'themeColors':
			return validateThemeColors(value, path)
		case 'by':
			return validateSortSpecs(value, path)
		case 'runs':
			return validateRichTextRuns(value, path)
		case 'style':
		case 'rule':
		case 'setup':
		case 'options':
		case 'protection':
		case 'properties':
		case 'documentProperties':
		case 'settings':
		case 'anchor':
			return isPlainObject(value) ? null : `${path} must be an object`
		case 'view':
			return value === null || isPlainObject(value) ? null : `${path} must be an object or null`
		default:
			return null
	}
}

function validateMode(op: unknown, value: unknown, path: string): string | null {
	if (op === 'setConditionalFormat') {
		return value === 'replace' || value === 'append'
			? null
			: `${path} must be one of replace, append`
	}
	if (
		op === 'setWorkbookProperties' ||
		op === 'setDocumentProperties' ||
		op === 'setWorkbookView'
	) {
		return value === 'merge' || value === 'replace' ? null : `${path} must be one of merge, replace`
	}
	if (op === 'copyRange' || op === 'moveRange') {
		return typeof value === 'string' && PASTE_MODES.has(value)
			? null
			: `${path} must be one of ${[...PASTE_MODES].join(', ')}`
	}
	return typeof value === 'string' ? null : `${path} must be a string`
}

function validateUpdates(value: unknown, path: string): string | null {
	if (!Array.isArray(value)) return `${path} must be an array`
	for (let i = 0; i < value.length; i++) {
		const update = value[i]
		if (!isPlainObject(update)) return `${path}[${i}] must be an object`
		if (typeof update.ref !== 'string') return `${path}[${i}].ref must be a string`
		if (!('value' in update)) return `${path}[${i}].value is required`
		if (!isInputValue(update.value)) return `${path}[${i}].value must be a scalar value or null`
	}
	return null
}

function validateRows(value: unknown, path: string): string | null {
	if (!Array.isArray(value)) return `${path} must be an array`
	for (let rowIndex = 0; rowIndex < value.length; rowIndex++) {
		const row = value[rowIndex]
		if (!Array.isArray(row)) return `${path}[${rowIndex}] must be an array`
		for (let colIndex = 0; colIndex < row.length; colIndex++) {
			if (!isInputValue(row[colIndex])) {
				return `${path}[${rowIndex}][${colIndex}] must be a scalar value or null`
			}
		}
	}
	return null
}

function validateStringArray(value: unknown, path: string): string | null {
	if (!Array.isArray(value)) return `${path} must be an array`
	for (let i = 0; i < value.length; i++) {
		if (typeof value[i] !== 'string') return `${path}[${i}] must be a string`
	}
	return null
}

function validateThemeColors(value: unknown, path: string): string | null {
	if (!Array.isArray(value)) return `${path} must be an array`
	for (let i = 0; i < value.length; i++) {
		const color = value[i]
		if (!isPlainObject(color)) return `${path}[${i}] must be an object`
		if (typeof color.slot !== 'string') return `${path}[${i}].slot must be a string`
		if ('rgb' in color && typeof color.rgb !== 'string') return `${path}[${i}].rgb must be a string`
		if ('systemColor' in color && typeof color.systemColor !== 'string') {
			return `${path}[${i}].systemColor must be a string`
		}
		if ('lastColor' in color && typeof color.lastColor !== 'string') {
			return `${path}[${i}].lastColor must be a string`
		}
	}
	return null
}

function validateSortSpecs(value: unknown, path: string): string | null {
	if (!Array.isArray(value)) return `${path} must be an array`
	for (let i = 0; i < value.length; i++) {
		const spec = value[i]
		if (!isPlainObject(spec)) return `${path}[${i}] must be an object`
		if (typeof spec.column !== 'string' && typeof spec.column !== 'number') {
			return `${path}[${i}].column must be a string or number`
		}
		if ('descending' in spec && typeof spec.descending !== 'boolean') {
			return `${path}[${i}].descending must be a boolean`
		}
	}
	return null
}

function validateRichTextRuns(value: unknown, path: string): string | null {
	if (!Array.isArray(value)) return `${path} must be an array`
	for (let i = 0; i < value.length; i++) {
		const run = value[i]
		if (!isPlainObject(run)) return `${path}[${i}] must be an object`
		if (typeof run.text !== 'string') return `${path}[${i}].text must be a string`
		for (const field of ['bold', 'italic', 'underline'] as const) {
			if (field in run && typeof run[field] !== 'boolean') {
				return `${path}[${i}].${field} must be a boolean`
			}
		}
		if ('color' in run && typeof run.color !== 'string')
			return `${path}[${i}].color must be a string`
		if ('size' in run && !isFiniteNumber(run.size)) return `${path}[${i}].size must be a number`
	}
	return null
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
	return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function isFiniteNumber(value: unknown): value is number {
	return typeof value === 'number' && Number.isFinite(value)
}

function isNonNegativeInteger(value: unknown): value is number {
	return Number.isInteger(value) && typeof value === 'number' && value >= 0
}

function isPositiveInteger(value: unknown): value is number {
	return Number.isInteger(value) && typeof value === 'number' && value > 0
}

function isInputValue(value: unknown): value is InputValue {
	return (
		value === null ||
		typeof value === 'string' ||
		typeof value === 'boolean' ||
		isFiniteNumber(value) ||
		value instanceof Date
	)
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
		case 'deleteDefinedName':
			return [
				'Review plan.approvals and provide --approval only for intentional deletion.',
				...common,
			]
		case 'insertRows':
		case 'insertCols':
			return [
				'Run ascend check first on imported workbooks; existing overlapping table ranges must be repaired before structural row or column edits.',
				'Confirm shifted table ranges remain non-overlapping after the insert.',
				...common,
			]
		case 'deleteRows':
			return [
				'Run ascend check first on imported workbooks; existing overlapping table ranges must be repaired before structural row edits.',
				'Do not delete only a table header or totals row; resize/delete the table explicitly, or delete the full table row span in one operation.',
				'Review plan.approvals and provide --approval only for intentional deletion.',
				...common,
			]
		case 'deleteCols':
			return [
				'Run ascend check first on imported workbooks; existing overlapping table ranges must be repaired before structural column edits.',
				'Rewrite or remove structured references to table fields before deleting the columns that contain those fields.',
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
			return [
				'Run ascend check first on imported workbooks; duplicate table names, duplicate ids, or overlapping table ranges must be repaired before table edits.',
				'Confirm the table range includes the intended header and data rows and does not overlap an existing table.',
				...common,
			]
		case 'appendRows':
			return [
				'Run ascend check first on imported workbooks; overlapping table ranges make append ownership ambiguous.',
				'Confirm appended rows will not expand the table into another table; totals-row appends can insert rows and must not shift another table.',
				...common,
			]
		case 'deleteTable':
			return [
				'Run ascend check first on imported workbooks; duplicate table names, duplicate ids, overlapping ranges, or structured-reference issues must be resolved before deleting table metadata.',
				'Rewrite or remove structured references that target the table before deleting its metadata.',
				'Review plan.approvals and provide --approval only for intentional deletion.',
				...common,
			]
		case 'renameTable':
			return [
				'Choose a workbook-unique table name; Excel treats table names case-insensitively in structured references.',
				'Use a name that starts with a letter, underscore, or backslash and is not C, R, A1-style, or R1C1-style.',
				'Run ascend check first on imported workbooks with duplicate table names or ids before renaming.',
				...common,
			]
		case 'setTableColumn':
			return [
				'Use a table column name or 0-based column index; set newName to rename the column and rewrite structured references.',
				'Set formula to null to clear calculated-column formulas.',
				'Use totalsRowFunction, totalsRowFormula, or totalsRowLabel to edit totals metadata.',
				...common,
			]
		case 'setTableStyle':
			return [
				'Use styleName to set or clear the table style name and booleans for first/last columns or row/column stripes.',
				...common,
			]
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
		case 'setDrawingText':
			return [
				'Use inspect --detail visuals or visualInventory to select drawingPartPath, id, name, or drawingObjectIndex.',
				'This edits existing shape/text-box text and preserves anchors, geometry, and drawing relationships.',
				...common,
			]
		case 'setThreadedComment':
			return [
				'Use inspectSheet threadedComments to choose threadedCommentId, partPath, ref, or commentIndex.',
				'This edits existing threaded comment text and preserves person, parent, timestamp, done, mention, and extension metadata.',
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
		case 'setPivotFieldItem':
			return [
				'Use inspect --detail pivots to choose pivotTable or partPath plus fieldIndex and itemIndex.',
				'Expect pivot output to be stale until Excel or another pivot-aware engine refreshes the pivot table.',
				...common,
			]
		case 'setSlicerCacheItem':
			return [
				'Use inspect --detail slicers to choose slicerCache or partPath and the zero-based item index.',
				'Expect pivot output to be stale until Excel or another pivot-aware engine refreshes the slicer-linked pivot tables.',
				...common,
			]
		case 'setTimelineRange':
			return [
				'Use inspect --detail timelines to choose timelineCache or partPath plus ISO-like startDate/endDate strings.',
				'Expect pivot output to be stale until Excel or another pivot-aware engine refreshes the timeline-linked pivot tables.',
				...common,
			]
		case 'setSparklineGroup':
			return [
				'Use inspectSheet().sparklineGroups to choose sheet and groupIndex.',
				'This rewrites preserved sparkline extension XML; keep source and location ranges shape-compatible.',
				...common,
			]
		case 'setAdvancedFilter':
			return [
				'Use inspectSheet().advancedFilters to choose sheet and filterIndex.',
				'This rewrites preserved custom sheet view autoFilter XML; use column+values for value-list filters.',
				...common,
			]
		case 'setConnectionRefresh':
			return [
				'Use inspect --detail connections or refreshMetadata to choose partPath, name, connectionId, or sheet.',
				'Set refreshOnLoad=true or saveData=false when external query output should be refreshed before use.',
				...common,
			]
		case 'rewriteExternalLink':
			return [
				'Use inspect --detail external-refs to choose partPath, relId, linkRelId, or target.',
				'Run ascend plan before commit to verify package preservation and formula-state warnings.',
				...common,
			]
		case 'setWorkbookProperties':
			return [
				'Use mode="merge" for targeted metadata edits or mode="replace" to rewrite workbookPr.',
				'Use null property values to clear codeName, defaultThemeVersion, filterPrivacy, or date1904.',
				...common,
			]
		case 'setDocumentProperties':
			return [
				'Use mode="merge" for targeted docProps edits or mode="replace" to rewrite core/app/custom document properties.',
				'Use null core/app fields to clear values; custom properties are replaced as a collection when provided.',
				...common,
			]
		case 'setWorkbookView':
			return [
				'Use index=0 for the primary workbook view; index equal to view count appends a new view.',
				'Use view=null to delete an existing view, or null field values to clear individual fields.',
				...common,
			]
		case 'setCalcSettings':
			return [
				'Use dateSystem carefully because changing it can alter date serial interpretation.',
				'Use iterativeCalc=null to disable iterative calculation and reset convergence defaults.',
				...common,
			]
		case 'setTheme':
			return [
				'Use inspect().themeSummary to preview existing theme names, fonts, and color slots.',
				'Theme color edits affect any style, chart, or drawing that references the edited theme slot.',
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
			return { op, sheet: 'Sheet1', range: 'A2:A20' }
		case 'setAutoFilter':
			return {
				op,
				sheet: 'Sheet1',
				range: 'A1:D20',
				column: 0,
				values: ['North'],
			}
		case 'setPrintArea':
			return { op, sheet: 'Sheet1', range: 'A1:D20' }
		case 'deleteConditionalFormat':
			return { op, sheet: 'Sheet1', range: 'A1:D20', priority: 1 }
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
			return {
				op,
				sheet: 'Sheet1',
				source: 'A1:B5',
				targetSheet: 'Summary',
				target: 'D1',
				mode: 'all',
			}
		case 'groupRows':
			return { op, sheet: 'Sheet1', from: 2, to: 10, collapsed: false }
		case 'groupCols':
			return { op, sheet: 'Sheet1', from: 2, to: 5, collapsed: false }
		case 'setRichText':
			return { op, sheet: 'Sheet1', ref: 'A1', runs: [{ text: 'Bold', font: { bold: true } }] }
		case 'setWorkbookProperties':
			return {
				op,
				properties: { codeName: 'Model', filterPrivacy: true, date1904: false },
				mode: 'merge',
			}
		case 'setDocumentProperties':
			return {
				op,
				properties: {
					core: { title: 'Forecast Pack', creator: 'Finance Ops' },
					app: { HeadingPairs: ['Worksheets', 1], TitlesOfParts: ['Sheet1'] },
					custom: [{ name: 'Reviewed', value: true }],
				},
				mode: 'merge',
			}
		case 'setWorkbookView':
			return { op, index: 0, view: { activeTab: 0, firstSheet: 0 }, mode: 'merge' }
		case 'setCalcSettings':
			return { op, settings: { calcMode: 'manual', fullCalcOnLoad: true } }
		case 'setTheme':
			return {
				op,
				themeName: 'Brand Theme',
				colorSchemeName: 'Brand Colors',
				majorFontLatin: 'Aptos Display',
				minorFontLatin: 'Aptos',
				themeColors: [{ slot: 'accent1', rgb: '0F6CBD' }],
			}
		case 'setWorkbookProtection':
			return { op, protection: { lockStructure: true } }
		case 'deleteTable':
			return { op, table: 'Sales' }
		case 'renameTable':
			return { op, table: 'Sales', newName: 'SalesData' }
		case 'resizeTable':
			return { op, table: 'Sales', ref: 'A1:E20' }
		case 'setTableColumn':
			return {
				op,
				table: 'Sales',
				column: 'Total',
				newName: 'Line Total',
				formula: '=[@Qty]*[@Price]',
				totalsRowFunction: 'sum',
			}
		case 'setTableStyle':
			return {
				op,
				table: 'Sales',
				styleName: 'TableStyleMedium2',
				showRowStripes: true,
			}
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
		case 'setDrawingText':
			return {
				op,
				sheet: 'Sheet1',
				drawingPartPath: 'xl/drawings/drawing1.xml',
				id: 2,
				text: 'Updated callout',
			}
		case 'setThreadedComment':
			return {
				op,
				sheet: 'Sheet1',
				threadedCommentId: '{thread-id}',
				text: 'Updated review note',
			}
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
		case 'setPivotFieldItem':
			return {
				op,
				pivotTable: 'PivotTable1',
				fieldIndex: 0,
				itemIndex: 2,
				hidden: true,
				selectedPageItem: 2,
			}
		case 'setSlicerCacheItem':
			return {
				op,
				slicerCache: 'Slicer_State',
				item: 0,
				selected: true,
				noData: false,
			}
		case 'setTimelineRange':
			return {
				op,
				timelineCache: 'Timeline_Order_Date',
				startDate: '2024-01-01T00:00:00',
				endDate: '2024-03-31T00:00:00',
			}
		case 'setSparklineGroup':
			return {
				op,
				sheet: 'Data',
				groupIndex: 0,
				range: 'Data!C2:C4',
				locationRange: 'E2:E4',
				type: 'column',
				markers: false,
			}
		case 'setAdvancedFilter':
			return {
				op,
				sheet: 'Data',
				filterIndex: 0,
				range: 'A1:D20',
				column: 0,
				values: ['East', 'North'],
				sortRef: 'A2:D20',
				sortBy: 'B2:B20',
				descending: false,
			}
		case 'setConnectionRefresh':
			return {
				op,
				partPath: 'xl/queryTables/queryTable1.xml',
				connectionId: 1,
				refreshOnLoad: true,
				saveData: false,
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
